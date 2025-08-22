%% 1. Input & Setup
try
    data = readtable('data_produksi_cleaned.xlsx');
    % PENTING: Ganti 'cons' dengan nama kolom yang benar jika berbeda.
    process_data = data.cons; 
    process_data = process_data(isfinite(process_data));
catch ME
    error('Gagal memuat atau menemukan kolom data. Error: %s', ME.message);
end

LSL = 0.1;
USL = 0.5;

%% 2. Jalankan Pipeline Analisis
stability_results = analyzeStability(process_data);
capability_results = analyzeCapability(process_data, LSL, USL, stability_results);

%% 3. Hasilkan Laporan Excel
output_filename = 'Laporan_Analisis_SPC_v2.1.xlsx';
generateExcelReport(output_filename, process_data, LSL, USL, stability_results, capability_results);
disp(['Laporan Excel "' output_filename '" telah berhasil dibuat dengan analisis yang lebih detail.']);

% FUNCTION DEFINITIONS
function results = analyzeStability(dataVector)
    results = struct();
    results.params = struct();
    violations_cell = {};

    n = numel(dataVector);
    if n < 2
        results.isStable = false; 
        results.reportText = 'GAGAL: Data tidak cukup.'; 
        return; 
    end
    
    CL = mean(dataVector);
    moving_ranges = abs(diff(dataVector));
    barMR = mean(moving_ranges);
    d2 = 1.128; % Konstanta statistik untuk subgroup n=2
    sigma = barMR / d2;
    UCL = CL + 3 * sigma;
    LCL = CL - 3 * sigma;
    
    results.params = struct('CL', CL, 'UCL', UCL, 'LCL', LCL, 'sigma', sigma, 'barMR', barMR);

    for i = 1:n
        if dataVector(i) > UCL
            keterangan = sprintf('Nilai %.3f melebihi UCL (%.3f)', dataVector(i), UCL);
            violations_cell = [violations_cell; {'Melebihi UCL', i, dataVector(i), keterangan}];
        elseif dataVector(i) < LCL
            keterangan = sprintf('Nilai %.3f di bawah LCL (%.3f)', dataVector(i), LCL);
            violations_cell = [violations_cell; {'Di Bawah LCL', i, dataVector(i), keterangan}];
        end
    end

    for i = 1:(n - 7)
        subset_idx = i:(i+7);
        subset_data = dataVector(subset_idx);
        if all(subset_data > CL) || all(subset_data < CL)
            violations_cell = [violations_cell; {'Run di Satu Sisi', sprintf('%d-%d', i, i+7), mean(subset_data), '8+ titik di satu sisi garis tengah.'}];
            break;
        end
    end

    for i = 1:(n - 5)
        subset_idx = i:(i+5);
        subset_data = dataVector(subset_idx);
        if all(diff(subset_data) > 0) || all(diff(subset_data) < 0)
            violations_cell = [violations_cell; {'Tren Naik/Turun', sprintf('%d-%d', i, i+5), mean(subset_data), '6+ titik terus naik atau turun.'}];
            break;
        end
    end
    
    if isempty(violations_cell)
        results.isStable = true;
        results.reportText = 'STABIL. Tidak ditemukan sinyal penyebab khusus.';
        results.violationsTable = table();
    else
        results.isStable = false;
        results.reportText = sprintf('TIDAK STABIL. Ditemukan total %d sinyal/pelanggaran.', size(violations_cell, 1));
        results.violationsTable = cell2table(violations_cell, 'VariableNames', {'Tipe_Pelanggaran', 'Indeks_Data', 'Nilai_Data_Aktual', 'Keterangan_Spesifik'});
    end
end


function results = analyzeCapability(dataVector, LSL, USL, stabilityStruct)
    results = struct();
    results.params = struct('Cp', NaN, 'Cpk', NaN);

    if ~stabilityStruct.isStable
        results.isApplicable = false;
        results.reportText = 'DITANGGUHKAN. Proses tidak stabil, analisis kapabilitas tidak valid.';
        return;
    end

    results.isApplicable = true;
    mu = stabilityStruct.params.CL;
    sigma = stabilityStruct.params.sigma;
    Cp = (USL - LSL) / (6 * sigma);
    Cpk = min((USL - mu) / (3 * sigma), (mu - LSL) / (3 * sigma));
    results.params.Cp = Cp;
    results.params.Cpk = Cpk;

    if Cpk >= 1.33
        perf_interp = 'KAPABEL.';
    elseif Cpk >= 1.00
        perf_interp = 'KAPABEL SECARA MARJINAL.';
    else
        perf_interp = 'TIDAK KAPABEL.';
    end

    if (Cp > 0) && (abs(Cp - Cpk) / Cp < 0.05)
        centering_interp = 'Pusat proses SANGAT BAIK.';
    else
        centering_interp = 'Pusat proses PERLU PENYESUAIAN.';
    end

    results.reportText = [perf_interp ' ' centering_interp];
end


function generateExcelReport(outputFilename, dataVector, LSL, USL, stability, capability)
    if exist(outputFilename, 'file'), delete(outputFilename); end

    summary_data = {
        'Laporan Analisis SPC', datestr(now, 'dd-mmm-yyyy HH:MM'); '', '';
        'Parameter Input', ''; 'Batas Spesifikasi Bawah (LSL)', LSL; 'Batas Spesifikasi Atas (USL)', USL; '', '';
        'Hasil Analisis Stabilitas', ''; 'Kesimpulan', stability.reportText; '', '';
        'Hasil Analisis Kapabilitas', ''; 'Kesimpulan', capability.reportText
    };
    writecell(summary_data, outputFilename, 'Sheet', 'Ringkasan Laporan');

    details_table = table( ...
        {'Statistik Dasar';'Statistik Dasar';'Batas Kendali';'Batas Kendali';'Batas Kendali';'Indeks Kapabilitas';'Indeks Kapabilitas'}, ...
        {'Rata-rata Proses (mu)';'Standar Deviasi (est.)';'Upper Control Limit (UCL)';'Center Line (CL)';'Lower Control Limit (LCL)';'Indeks Cp';'Indeks Cpk'}, ...
        [stability.params.CL; stability.params.sigma; stability.params.UCL; stability.params.CL; stability.params.LCL; capability.params.Cp; capability.params.Cpk], ...
        {'Pusat data proses.';'Ukuran sebaran.';'Batas atas variasi alami.';'Sama dengan rata-rata.';'Batas bawah variasi alami.';'Potensi kapabilitas.';'Kinerja kapabilitas aktual.'}, ...
        'VariableNames', {'Kategori', 'Metrik', 'Nilai', 'Keterangan_Singkat'});
    writetable(details_table, outputFilename, 'Sheet', 'Detail Perhitungan');

    if ~stability.isStable && ~isempty(stability.violationsTable)
        writetable(stability.violationsTable, outputFilename, 'Sheet', 'Data Pelanggaran (Stabilitas)');
    end

    if ~stability.isStable && ~isempty(stability.violationsTable)
        above_ucl_violations = stability.violationsTable(strcmp(stability.violationsTable.Tipe_Pelanggaran, 'Melebihi UCL'), :);
        if ~isempty(above_ucl_violations)
            violation_values = above_ucl_violations.Nilai_Data_Aktual;
            bin_width = 0.1;
            max_val = max(violation_values);
            bin_edges = stability.params.UCL:bin_width:(ceil(max_val/bin_width)*bin_width);
            if isempty(bin_edges) || bin_edges(end) < max_val, bin_edges(end+1) = bin_edges(end) + bin_width; end
            [counts, ~] = histcounts(violation_values, bin_edges);
            total_violations = sum(counts);
            if total_violations > 0
                bin_labels = cell(length(bin_edges)-1, 1);
                for i = 1:length(bin_edges)-1, bin_labels{i,1} = sprintf('%.3f - %.3f', bin_edges(i), bin_edges(i+1)); end
                percentages = (counts' / total_violations) * 100;
                analysis_table = table(bin_labels, counts', percentages, 'VariableNames', {'Rentang_Konsentrasi_UCL', 'Jumlah_Pelanggaran', 'Persentase'});
                total_row = table({'Total'}, total_violations, 100, 'VariableNames', {'Rentang_Konsentrasi_UCL', 'Jumlah_Pelanggaran', 'Persentase'});
                analysis_table = [analysis_table; total_row];
                writetable(analysis_table, outputFilename, 'Sheet', 'Analisis Pelanggaran Atas (UCL)');
            end
        end
    end
    
    spec_violations_cell = {};
    for i = 1:numel(dataVector)
        if dataVector(i) > USL
            keterangan = sprintf('Nilai %.3f melebihi Batas Atas Spesifikasi (USL = %.3f)', dataVector(i), USL);
            spec_violations_cell = [spec_violations_cell; {'Di Atas USL', i, dataVector(i), keterangan}];
        elseif dataVector(i) < LSL
            keterangan = sprintf('Nilai %.3f di bawah Batas Bawah Spesifikasi (LSL = %.3f)', dataVector(i), LSL);
            spec_violations_cell = [spec_violations_cell; {'Di Bawah LSL', i, dataVector(i), keterangan}];
        end
    end
    
    if ~isempty(spec_violations_cell)
        spec_violations_table = cell2table(spec_violations_cell, 'VariableNames', {'Tipe_Pelanggaran_Spesifikasi', 'Indeks_Data', 'Nilai_Data_Aktual', 'Keterangan_Spesifik'});
        writetable(spec_violations_table, outputFilename, 'Sheet', 'Analisis Luar Spesifikasi', 'Range', 'A1');
        
        above_usl_data = dataVector(dataVector > USL);
        if ~isempty(above_usl_data)
            bin_width = 0.1;
            max_val = max(above_usl_data);
            bin_edges = USL:bin_width:(ceil(max_val/bin_width)*bin_width);
            if isempty(bin_edges) || bin_edges(end) < max_val, bin_edges(end+1) = bin_edges(end) + bin_width; end
            
            [counts, ~] = histcounts(above_usl_data, bin_edges);
            total_points = sum(counts);
            
            if total_points > 0
                bin_labels = cell(length(bin_edges)-1, 1);
                for i = 1:length(bin_edges)-1, bin_labels{i,1} = sprintf('%.1f - %.1f', bin_edges(i), bin_edges(i+1)); end
                percentages = (counts' / total_points) * 100;
                
                dist_table = table(bin_labels, counts', percentages, 'VariableNames', {'Rentang_Konsentrasi_USL', 'Jumlah_Data', 'Persentase'});
                total_row = table({'Total'}, total_points, 100, 'VariableNames', {'Rentang_Konsentrasi_USL', 'Jumlah_Data', 'Persentase'});
                dist_table = [dist_table; total_row];
                
                writetable(dist_table, outputFilename, 'Sheet', 'Analisis Luar Spesifikasi', 'Range', 'F1');
            end
        end
    end
end
