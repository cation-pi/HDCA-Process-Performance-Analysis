% PROCESS STABILITY AND CAPABILITY ANALYSIS
% HDCA Production Performance Evaluation
% CTQ: Final Residual Hydrazine Hydrate Concentration (cons)

filename = 'data_produksi_cleaned.xlsx';
variable_to_analyze = 'cons';

% Define specification limits (LSL & USL) for product acceptance 
LSL = 0.1;
USL = 0.5;
fprintf('--- Starting SPC Analysis for CTQ: %s ---\n\n', variable_to_analyze);

disp('Part 1: PROCESS STABILITY ANALYSIS');
try
    dataTable = readtable(filename);
    data = dataTable.(variable_to_analyze);
    data = data(isfinite(data));
    fprintf('Successfully loaded and cleaned data from "%s".\n', filename);
catch ME
    error('Failed to load or process the file "%s". Original error: %s', filename, variable_to_analyze, ME.message);
end

fprintf('Generating I-MR control chart...\n');
figure('Name', 'Figure 1: I-MR Control Chart', 'NumberTitle', 'off', 'WindowState', 'maximized');
center_line = mean(data);
moving_range = abs(diff(data));
avg_moving_range = mean(moving_range);

% d2 constant for a subgroup size of 2 (for moving range) is 1.128
sigma_est = avg_moving_range / 1.128; 
UCL = center_line + 3 * sigma_est;
LCL = center_line - 3 * sigma_est;

ax1 = subplot(2,1,1); 
plot(data, '-o', 'MarkerFaceColor', 'b', 'MarkerEdgeColor', 'b');
hold on;
line([1, numel(data)], [center_line, center_line], 'Color', 'k', 'LineStyle', '-', 'LineWidth', 1.5);
line([1, numel(data)], [UCL, UCL], 'Color', 'r', 'LineStyle', '--', 'LineWidth', 2.5);
line([1, numel(data)], [LCL, LCL], 'Color', 'r', 'LineStyle', '--', 'LineWidth', 2.5);
title('Individual Chart untuk Konsentrasi Sisa H/H pada Proses HDCA', 'FontSize', 16);
ylabel('Individual Value');
grid on;

ax2 = subplot(2,1,2);
plot(moving_range, '-o', 'MarkerFaceColor', 'b', 'MarkerEdgeColor', 'b');
hold on;
mr_UCL = 3.267 * avg_moving_range; % D4 constant for n=2
line([1, numel(moving_range)], [avg_moving_range, avg_moving_range], 'Color', 'k', 'LineStyle', '-', 'LineWidth', 1.5);
line([1, numel(moving_range)], [mr_UCL, mr_UCL], 'Color', 'r', 'LineStyle', '--', 'LineWidth', 2.5);
title('Moving Range Chart', 'FontSize', 16);
xlabel('Batch Number');
ylabel('Moving Range');
grid on;

linkaxes([ax1, ax2], 'x');
axis(ax1, 'tight');

fprintf('Applying standard control rules (Nelson Rules)...\n');
violating_points = [];

rule1_violations = find(data > UCL | data < LCL);
violating_points = [violating_points; rule1_violations];

for i = 8:numel(data)
    if all(data(i-7:i) > center_line) || all(data(i-7:i) < center_line)
        violating_points = [violating_points; (i-7:i)'];
    end
end

for i = 6:numel(data)
    subset = data(i-5:i);
    if all(diff(subset) > 0) || all(diff(subset) < 0)
        violating_points = [violating_points; (i-5:i)'];
    end
end

violating_points = unique(violating_points);
if ~isempty(violating_points)
    subplot(2,1,1); % Switch back to the Individuals chart
    plot(violating_points, data(violating_points), 'rs', 'MarkerSize', 5, 'MarkerFaceColor', 'r');
end

saveas(gcf, 'I-MR_Control_Chart.png');
fprintf('I-MR chart saved as I-MR_Control_Chart.png\n');

num_violations = numel(violating_points);
if num_violations == 0
    stability_status = 'STABLE';
    fprintf('STABILITY CONCLUSION: PROCESS IS STABLE.\n');
    disp('No out-of-control signals were detected.');
    disp('Proceeding to Capability Analysis.');
else
    stability_status = 'UNSTABLE';
    fprintf('STABILITY CONCLUSION: PROCESS IS UNSTABLE.\n');
    fprintf('[%d] out-of-control signal(s) detected.\n', num_violations);
    disp('WARNING: Capability analysis results may not be reliable for an unstable process.');
end
fprintf('\n');

disp('Part 2: PROCESS CAPABILITY ANALYSIS');

% This section runs regardless of stability but includes a warning if unstable.
if strcmp(stability_status, 'UNSTABLE')
    disp('Executing capability analysis with a note of caution due to process instability.');
end

fprintf('Generating process capability report...\n');
figure('Name', 'Figure 2: Process Capability Histogram', 'NumberTitle', 'off', 'WindowState', 'maximized');

histogram(data, 'Normalization', 'pdf', 'FaceColor', [0.5 0.7 1.0]);
hold on;
x_range = linspace(min(data), max(data), 100);
pdf_curve = normpdf(x_range, center_line, sigma_est);
plot(x_range, pdf_curve, 'k-', 'LineWidth', 2);

line([LSL, LSL], ylim, 'Color', 'r', 'LineStyle', '--', 'LineWidth', 2.5);
line([USL, USL], ylim, 'Color', 'r', 'LineStyle', '--', 'LineWidth', 2.5);
line([center_line, center_line], ylim, 'Color', [0 0.5 0], 'LineStyle', '-', 'LineWidth', 2.5);
legend('Data', 'Normal Distribution Fit', 'LSL', 'USL', 'Process Mean');
xlabel('Konsentrasi Sisa H/H');
ylabel('Densitas');
grid on;
title('Plot Histogram untuk Analisis Kapabilitas Proses Produksi HDCA Berdasarkan Konsentrasi Sisa H/H', 'FontSize', 16);

saveas(gcf, 'Process_Capability_Histogram.png');
fprintf('Process capability plot saved as Process_Capability_Histogram.png\n\n');

Cp = (USL - LSL) / (6 * sigma_est);
Cpu = (USL - center_line) / (3 * sigma_est);
Cpl = (center_line - LSL) / (3 * sigma_est);
Cpk = min(Cpu, Cpl);

fprintf('Calculated Capability Indices:\n');
fprintf('   Cp:  %.3f\n', Cp);
fprintf('   Cpk: %.3f\n\n', Cpk);

disp('PERFORMANCE INTERPRETATION:');
if Cpk >= 1.67
    disp('   EXCELLENT. Process is considered world-class (Six Sigma level).');
elseif Cpk >= 1.33
    disp('   CAPABLE. Process meets the standard minimum for most industries.');
elseif Cpk >= 1.00
    disp('   MARGINALLY CAPABLE. Process is susceptible to producing defects if small shifts occur.');
else
    disp('   NOT CAPABLE. Process is producing products outside of the specification limits.');
end
fprintf('\n');

disp('CENTERING ANALYSIS:');
% Check if the difference is significant (e.g., more than 10-15% of Cp)
if abs(Cp - Cpk) < 0.15 * Cp
    disp('   EXCELLENT. The process mean is well-centered between the specification limits.');
else
    disp('   IMPROVEMENT NEEDED. The process variation is low, but the mean is off-center.');
end

fprintf('\n--- Analysis Complete ---\n');
