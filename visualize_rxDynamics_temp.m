% This script focuses on visualizing the reaction dynamics of reactor temperature over time. 
% Its purpose is to compare temperature profiles across different batches and to display the average reaction kinetics profile from the entire production dataset.

inputFile = 'DataHDCA_TempxTime.xlsx';
folderPlots = 'Plots_HDCA';
outputFileName = fullfile(folderPlots, 'Profil_Waktu_vs_Suhu_Batch.png');

if ~isfolder(folderPlots)
    mkdir(folderPlots);
    fprintf('Folder output "%s" telah dibuat.\n', folderPlots);
end

try
    batchNames = sheetnames(inputFile);
    fprintf('Berhasil membaca file: "%s"\n', inputFile);
    fprintf('Jumlah batch yang terdeteksi: %d\n\n', numel(batchNames));
catch ME
    error('Gagal membaca file Excel: "%s". Pastikan file ada di path yang benar.\nError: %s', inputFile, ME.message);
end

fig = figure('Name', 'Profil Waktu vs. Suhu per Batch', 'Position', [100, 100, 1100, 700]);
hold on;
grid on;
legendEntries = {};

fprintf('Memulai proses plotting untuk setiap batch...\n');
for i = 1:numel(batchNames)
    currentSheetName = batchNames(i);
    fprintf('  - Memproses Batch: %s\n', currentSheetName);
    
    data = readtable(inputFile, 'Sheet', currentSheetName);
    
    data = rmmissing(data, 'DataVariables', {'temp', 'jam_ke'});
    
    plot(data.jam_ke, data.temp, 'o-', 'LineWidth', 1.5, 'MarkerSize', 5);
    
    legendEntries{end+1} = currentSheetName;
end
fprintf('Semua batch berhasil diplot.\n\n');

hold off;
legend(legendEntries, 'Location', 'best', 'Interpreter', 'none');
title('Perbandingan Profil Waktu Reaksi vs. Suhu Reaktor Antar Batch', 'FontSize', 16, 'FontWeight', 'bold');
xlabel('Waktu Reaksi Kumulatif (Jam)', 'FontSize', 12, 'FontWeight', 'bold');
ylabel('Suhu Reaktor (°C)', 'FontSize', 12, 'FontWeight', 'bold');

ax = gca;
ax.FontSize = 11;
ax.Box = 'on';

try
    print(fig, outputFileName, '-dpng', '-r300');
    fprintf('PROSES SELESAI.\n');
    fprintf('Plot berhasil disimpan sebagai: %s\n', outputFileName);
catch ME
    fprintf('ERROR: Gagal menyimpan file gambar.\n');
    fprintf('Pastikan Anda memiliki izin tulis di folder target.\n');
    fprintf('Error: %s\n', ME.message);
end
