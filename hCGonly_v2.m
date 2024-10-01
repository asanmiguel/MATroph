clc
clear all
close all

%% Define image directories and sheet name
dirname_c = "D:\STB hCG Quantitative Analysis\Isotype Control\"; % Control image folder
dirnameconstant = "D:\STB hCG Quantitative Analysis\Experimental"; % Experimental folder base
dirnamevariable = ["condition 1","condition 2"]; % Subdirectories for different conditions
sheetname = 'STB analysis.xlsx'; % Output Excel file

%% Process control images
imagelist = dir(fullfile(dirname_c, '*.tif')); % Identify .tif files in control folder
total_objects_Blue_c = 0; % Initialize total for control images

for i = 1:length(imagelist)
    disp('Working on Control Image')

    file_c = strcat(dirname_c, imagelist(i).name); % Full path to control image
    im_c = imread(file_c); % Read the control image

    % Process blue channel
    [Blueim_c, Blue_c] = process_r_g(im_c, 3);

    % Accumulate total number of blue objects
    total_objects_Blue_c = total_objects_Blue_c + numel(Blue_c);

    % Process red channel and other operations 
    [RedSum_c(:,i), Red_mean_c(:,i)] = hCGintensity2Dfunctionbacksubtracted(im_c, 1);
end

%% Process experimental conditions
for x = 1:length(dirnamevariable)
    total_objects_Blue = 0; % Initialize total for experimental images

    dirnames = dirnameconstant + "\" + dirnamevariable(x) + "\"; % Directory for current condition
    dirname = dirnames{1};
    disp(dirname)

    images = dir(fullfile(dirname, '*.tif')); % Identify .tif files in experimental folder

    for i = 1:length(images)
        disp('Working on Experimental Image')

        file = strcat(dirname, images(i).name); % Full path to experimental image
        im = imread(file); % Read the experimental image

        % Process blue channel
        [Blueim, Blue] = process_r_g(im, 3);

        % Accumulate total number of blue objects
        total_objects_Blue = total_objects_Blue + numel(Blue);

        % Process red channel and other operations 
        [RedSum(:,i), Red_mean(:,i)] = hCGintensity2Dfunctionbacksubtracted(im, 1);
    end

    % Get the condition name for the current sheet
    v = dirnamevariable(x);

    %% Output results to Excel
    disp("printing to " + sheetname)

    % Write Red_Sum and Red_Mean for experimental condition
    writematrix("Red_Sum", sheetname, 'Sheet', v, 'Range', "A1", 'UseExcel', true)
    writematrix(RedSum(:), sheetname, 'Sheet', v, 'Range', "A2:A" + (length(RedSum) + 1), 'UseExcel', true)
    writematrix("Red_Mean", sheetname, 'Sheet', v, 'Range', "B1", 'UseExcel', true)
    writematrix(Red_mean(:), sheetname, 'Sheet', v, 'Range', "B2:B" + (length(Red_mean) + 1), 'UseExcel', true)

    % Write RedSum_c and Red_mean_c for control condition
    writematrix("RedSum_c", sheetname, 'Sheet', v, 'Range', "C1", 'UseExcel', true)
    writematrix(RedSum_c(:), sheetname, 'Sheet', v, 'Range', "C2:C" + (length(RedSum_c) + 1), 'UseExcel', true)
    writematrix("Red_mean_c", sheetname, 'Sheet', v, 'Range', "D1", 'UseExcel', true)
    writematrix(Red_mean_c(:), sheetname, 'Sheet', v, 'Range', "D2:D" + (length(Red_mean_c) + 1), 'UseExcel', true)

    % Write total number of blue objects for both control and experimental
    writematrix("Total Objects Blue_c", sheetname, 'Sheet', v, 'Range', "F1", 'UseExcel', true)
    writematrix(total_objects_Blue_c, sheetname, 'Sheet', v, 'Range', "F2", 'UseExcel', true)

    writematrix("Total Objects Blue", sheetname, 'Sheet', v, 'Range', "E1", 'UseExcel', true)
    writematrix(total_objects_Blue, sheetname, 'Sheet', v, 'Range', "E2", 'UseExcel', true)
end
