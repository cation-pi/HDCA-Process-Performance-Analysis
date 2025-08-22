% INTERPRETASI SPC SECARA OTOMATIS

namaFileMasukan = 'data_produksi_cleaned.xlsx';
namaFileKeluaran = 'Laporan_Kinerja_HDCA.xlsx';

try
    T = readtable(namaFileMasukan);
    disp('Data berhasil dibaca dari file CSV.');
catch ME
    error('File "%s" tidak ditemukan');
end

tanggalDalamTeks = T.date;
tanggalDalamDatetime = NaT(size(tanggalDalamTeks));
for i = 1:length(tanggalDalamTeks)
    teks = tanggalDalamTeks{i};
    if ~isempty(teks) && ~all(isspace(teks))
        try
            tanggalDalamDatetime(i) = datetime(teks, 'InputFormat', 'dd/MM/yy');
        catch
            fprintf(2, 'Peringatan: Format tanggal salah di baris Excel %d (nilai: "%s"). Baris ini akan diabaikan.\n', i+1, teks);
        end
    end
end
T.date = tanggalDalamDatetime;
T = rmmissing(T, 'DataVariables', {'date'});
T.Bulan = month(T.date);

T = T(T.Bulan ~= 8, :);
disp('Data untuk bulan Agustus telah dikecualikan dari analisis.');
bulanUnik = unique(T.Bulan);
jumlahBulan = numel(bulanUnik);
hasilAnalisis = table();

disp('Menghitung rata-rata kinerja bulanan...');
for i = 1:jumlahBulan
    bulanIni = bulanUnik(i);
    dataBulanIni = T(T.Bulan == bulanIni, :);
    
    tempHasil.Bulan = bulanIni;
    tempHasil.Massa_H_H_kg = mean(dataBulanIni.h_h_kg_1, 'omitnan');
    tempHasil.Massa_Urea_kg = mean(dataBulanIni.urea_kg_1, 'omitnan');
    tempHasil.HH_Charge = mean(dataBulanIni.hh_charge_1, 'omitnan');
    tempHasil.Dissolver_HH = mean(dataBulanIni.dissolver_hh__1, 'omitnan');
    tempHasil.FQ_Awal = mean(dataBulanIni.fq_awal_2, 'omitnan');
    tempHasil.FQ_Akhir = mean(dataBulanIni.fq_akhir_2, 'omitnan');
    tempHasil.Total_Charging_kg = sum(dataBulanIni.total_charge_kg_2, 'omitnan');
    tempHasil.Konsentrasi_Awal_HH = mean(dataBulanIni.cons_hh_2, 'omitnan');
    tempHasil.Total_Waktu_Reaksi = mean(dataBulanIni.total_jam, 'omitnan');
    tempHasil.Suhu_Awal_Reaktor = mean(dataBulanIni.temp_awal, 'omitnan');
    tempHasil.Suhu_Akhir_Reaktor = mean(dataBulanIni.temp_akhir, 'omitnan');
    tempHasil.Konsentrasi_Sisa_HH = mean(dataBulanIni.cons, 'omitnan');
    tempHasil.Hasil_Produksi_kg = sum(dataBulanIni.hasil_produksi_kg, 'omitnan');
    
    hasilAnalisis = [hasilAnalisis; struct2table(tempHasil, 'AsArray', true)];
end

disp('Menghitung rata-rata kinerja keseluruhan...');
overall_table = table(...
    {'Overall'}, ...
    mean(T.h_h_kg_1, 'omitnan'), ...
    mean(T.urea_kg_1, 'omitnan'), ...
    mean(T.hh_charge_1, 'omitnan'), ...
    mean(T.dissolver_hh__1, 'omitnan'), ...
    mean(T.fq_awal_2, 'omitnan'), ...
    mean(T.fq_akhir_2, 'omitnan'), ...
    mean(T.total_charge_kg_2, 'omitnan'), ...
    mean(T.cons_hh_2, 'omitnan'), ...
    mean(T.total_jam, 'omitnan'), ...
    mean(T.temp_awal, 'omitnan'), ...
    mean(T.temp_akhir, 'omitnan'), ...
    mean(T.cons, 'omitnan'), ...
    mean(T.hasil_produksi_kg, 'omitnan'), ...
    'VariableNames', hasilAnalisis.Properties.VariableNames);
hasilAnalisis.Bulan = cellstr(num2str(hasilAnalisis.Bulan));
hasilAnalisis = [hasilAnalisis; overall_table];
hasilAnalisis.Properties.VariableNames = {
    'Bulan', 'Massa H/H (kg)', 'Massa Urea (kg)', 'H/H% Charging', ...
    'Dissolver H/H%', 'FQ Awal', 'FQ Akhir', 'Total Charging (kg)', ...
    'Konsentrasi Awal H/H (%)', 'Total Waktu Reaksi', 'Suhu Awal Reaktor', ...
    'Suhu Akhir Reaktor', 'Konsentrasi Sisa H/H (%)', 'Hasil Produksi /batch'
};

try
    writetable(hasilAnalisis, namaFileKeluaran, 'Sheet', 1, 'WriteRowNames',false);
    fprintf('Laporan berhasil disimpan sebagai "%s".\n', namaFileKeluaran);
catch ME
    error('Gagal menyimpan file Excel. Error: %s', ME.message);
end

disp('Membuat visualisasi data final...');
plotData = hasilAnalisis(1:end-1,:);
bulanNumerik = str2double(plotData.Bulan);
bulanDatetime = datetime(2024, bulanNumerik, 1);
namaBulan = cellstr(datestr(bulanDatetime, 'mmmm'));
fontSizeBesar = 10;

% Visualisasi 1
figure('Name', 'Produksi vs Charging', 'NumberTitle', 'off');
barData = [plotData.('Hasil Produksi /batch'), plotData.('Total Charging (kg)')];
b = bar(barData, 'grouped');
hold on;
for i = 1:length(b)
    xPos = b(i).XEndPoints;
    yPos = b(i).YEndPoints;
    labels = string(round(yPos));
    text(xPos, yPos, labels, ...
         'HorizontalAlignment', 'left', ...
         'VerticalAlignment', 'middle', ...
         'Rotation', 45, ...
         'FontSize', fontSizeBesar);
end
hold off;
set(gca, 'XTickLabel', namaBulan);
title('Perbandingan Total Produksi vs Total Charging per Bulan');
xlabel('Bulan');
ylabel('Jumlah (kg)');
legend({'Total Produksi', 'Total Charging'}, 'Location', 'northoutside', 'Orientation', 'horizontal');
grid on;
ylim([0, max(barData(:)) * 1.2]); 
saveas(gcf, '1_Produksi_vs_Charging_final.png');

% Visualisasi 2
figure('Name', 'Parameter Operasional Kunci', 'NumberTitle', 'off');
x_coords = 1:jumlahBulan;
yyaxis left;
y_suhu = plotData.('Suhu Akhir Reaktor');
plot(x_coords, y_suhu, '-o', 'LineWidth', 1.5);
ylabel('Suhu Akhir Reaktor (°C)');
hold on;
for i = 1:length(x_coords)
    text(x_coords(i), y_suhu(i), sprintf(' %.1f', y_suhu(i)), 'FontSize', fontSizeBesar, 'VerticalAlignment', 'bottom');
end
current_ylim = ylim;
ylim([current_ylim(1), current_ylim(2) * 1.05]);
yyaxis right;
y_konsentrasi = plotData.('Konsentrasi Sisa H/H (%)');
plot(x_coords, y_konsentrasi, '-s', 'LineWidth', 1.5);
ylabel('Konsentrasi Sisa H/H (%)');
for i = 1:length(x_coords)
    text(x_coords(i), y_konsentrasi(i), sprintf(' %.2f', y_konsentrasi(i)), 'FontSize', fontSizeBesar, 'VerticalAlignment', 'top');
end
current_ylim = ylim;
ylim([current_ylim(1) - current_ylim(2)*0.05, current_ylim(2)]);
hold off;
set(gca, 'XTick', x_coords, 'XTickLabel', namaBulan);
title('Analisis Parameter Operasional Kunci per Bulan');
xlabel('Bulan');
legend({'Suhu Akhir', 'Konsentrasi Sisa H/H'}, 'Location', 'best');
grid on;
saveas(gcf, '2_Parameter_Operasional_final_v2.png');

% Visualisasi 3
figure('Name', 'Efisiensi Produksi', 'NumberTitle', 'off');
efisiensi = plotData.('Hasil Produksi /batch') ./ plotData.('Total Charging (kg)');
bar(efisiensi);
hold on;
for i = 1:length(efisiensi)
    text(i, efisiensi(i), sprintf('%.2f', efisiensi(i)), ...
         'HorizontalAlignment','center', ...
         'VerticalAlignment','bottom', ...
         'FontSize', 25);
end
overallEfficiency = mean(T.hasil_produksi_kg, 'omitnan') / mean(T.total_charge_kg_2, 'omitnan');
yline(overallEfficiency, '--r', ['Rata-rata Efisiensi: ' num2str(overallEfficiency, '%.2f')], 'LineWidth', 3, 'FontSize', 25);
hold off;
set(gca, 'XTickLabel', namaBulan);
title('Analisis Efisiensi Produksi per Bulan (Hasil/Charging)');
xlabel('Bulan');
ylabel('Rasio Efisiensi');
grid on;
if ~isempty(efisiensi) && all(isfinite(efisiensi))
    ylim([min(efisiensi)*0.85, max(efisiensi)*1.15]);
end
exportgraphics(gcf, '3_Efisiensi_Produksi_final.png','Resolution',300);

disp('Selesai');
