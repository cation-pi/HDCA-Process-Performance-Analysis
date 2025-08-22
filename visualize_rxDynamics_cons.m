% This script focuses on visualizing the reaction dynamics of residual H/H concentration over time. 
% Its purpose is to display the average reaction kinetics profile from the entire production dataset.

filename = 'DataHDCA_TempxTime.xlsx';

if ~isfile(filename)
    error('File Excel "%s" tidak ditemukan di direktori saat ini.', filename);
end

try
    sheet_names = sheetnames(filename);
catch ME
    error('Gagal membaca nama sheet dari file. Pastikan file tidak rusak dan formatnya benar (.xlsx). Error: %s', ME.message);
end

fprintf('Menemukan %d sheet (batch) untuk digabungkan dari file "%s".\n', length(sheet_names), filename);

all_batch_data = {};
max_time = 0;
for i = 1:length(sheet_names)
    current_sheet = sheet_names{i};
    
    try
        opts = detectImportOptions(filename, 'Sheet', current_sheet, 'VariableNamingRule', 'preserve');
        dataTable = readtable(filename, opts);
    catch ME
        warning('Gagal membaca data dari sheet: "%s". Melewati sheet ini. Error: %s', current_sheet, ME.message);
        continue;
    end
    
    if ~ismember('jam_ke', dataTable.Properties.VariableNames) || ~ismember('cons_%', dataTable.Properties.VariableNames)
        warning('Sheet "%s" tidak memiliki kolom "jam_ke" atau "cons_%%". Melewati sheet ini.', current_sheet);
        continue;
    end
    
    time_hours = dataTable.jam_ke;
    concentration = dataTable.('cons_%');
    
    valid_indices = ~isnan(concentration);
    time_filtered = time_hours(valid_indices);
    concentration_filtered = concentration(valid_indices);
    
    if ~isempty(time_filtered)
        all_batch_data{end+1} = [time_filtered, concentration_filtered];
        max_time = max(max_time, max(time_filtered));
    end
end

fprintf('Berhasil memproses dan menyimpan data dari %d batch.\n', length(all_batch_data));

common_time_vector = (0:1:ceil(max_time))';
num_batches = length(all_batch_data);
num_time_points = length(common_time_vector);

interpolated_concentrations = NaN(num_time_points, num_batches);

for i = 1:num_batches
    batch_time = all_batch_data{i}(:, 1);
    batch_conc = all_batch_data{i}(:, 2);
    
    interpolated_concentrations(:, i) = interp1(batch_time, batch_conc, common_time_vector, 'linear', 'extrap');
end

interpolated_concentrations(interpolated_concentrations < 0) = 0;

mean_concentration = mean(interpolated_concentrations, 2, 'omitnan');
std_concentration = std(interpolated_concentrations, 0, 2, 'omitnan');

upper_bound = mean_concentration + std_concentration;
lower_bound = mean_concentration - std_concentration;
lower_bound(lower_bound < 0) = 0;

figure('Name', 'Average Reaction Kinetics Profile', 'NumberTitle', 'off');
hold on;

fill([common_time_vector', fliplr(common_time_vector')], [upper_bound', fliplr(lower_bound')], ...
     [0.8 0.8 1], 'EdgeColor', 'none', 'FaceAlpha', 0.5);

plot(common_time_vector, mean_concentration, 'b-', 'LineWidth', 2);

hold off;

title('Profil Rata-rata Penurunan Konsentrasi Hidrazin Hidrat', 'FontSize', 16, 'FontWeight', 'bold');
xlabel('Waktu Reaksi (jam)', 'FontSize', 12);
ylabel('Konsentrasi Sisa Hidrazin Hidrat (%)', 'FontSize', 12);
legend({'Variasi Proses (±1 Std Dev)', 'Profil Rata-rata'}, 'Location', 'northeast');
grid on;
box on;
axis tight;

current_ylimits = ylim(gca);
ylim(gca, [0, current_ylimits(2) * 1.1]);

output_filename = 'Profil_Kinetika_Rata_Rata.png';
try
    print(gcf, output_filename, '-dpng', '-r300');
    fprintf('\nPlot telah berhasil disimpan sebagai "%s".\n', output_filename);
catch ME
    warning('Gagal menyimpan plot. Error: %s', ME.message);
end

fprintf('Plotting complete. Menampilkan profil rata-rata dari %d batch.\n', num_batches);
