% This section of script gives the sum of CumInjWater_bbl_ for each
% SortUWI within a specified date range
% Replace 'your_file.xlsx' with the actual name of your Excel file
%excelFileName = 'Injection_History.xlsx';
%excelFileName = 'Production_History.xlsx';
excelFileName = 'Horizontals_Production_History.xlsx';

% Read the Excel file into a table
dataTable = readtable(excelFileName);

% Extract the column with the title 'Sort UWI' and 'ProdDate'
sortUWIColumn = dataTable.('SortUWI');
prodDateColumn = dataTable.('ProdDate');

% Remove empty or 0x0 character values
sortUWIColumn = sortUWIColumn(~cellfun('isempty', sortUWIColumn));

% Count unique values in the 'Sort UWI' column
uniqueValues = unique(sortUWIColumn);

% Specify the date range
startDate = datetime('01-Mar-2011', 'Format', 'dd-MMM-yyyy');
endDate = datetime('01-Jun-2023', 'Format', 'dd-MMM-yyyy');

% Create a cell array to store the results
resultsCellArray = {'Date', 'Total AvgDlyOil_bbl_d_'};

% Loop through each unique date in the specified range
dateRange = startDate:endDate;
for d = 1:length(dateRange)
    currentDate = dateRange(d);
    
    % Filter data for the current date
    idxDate = prodDateColumn == currentDate;
    
    % Get the sum of CumInjWater_bbl_ for all SortUWIs on the current date
    totalCumm = sum(dataTable.AvgDlyOil_bbl_d_(idxDate));
    
    % Display and store the results
    disp(['Date: ', datestr(currentDate, 'dd-mmm-yyyy'), ', Total AvgDlyOil_bbl_d_ for all SortUWIs: ', num2str(totalCumm)]);
    
    % Store results in the cell array
    resultsCellArray = [resultsCellArray; {currentDate, totalCumm}];
end

% Convert the cell array to a table
resultsTable = cell2table(resultsCellArray(2:end, :), 'VariableNames', resultsCellArray(1, :));

% Remove rows with 0 values in 'Total CumInjWater_bbl_' column
resultsTable(resultsTable.('Total AvgDlyOil_bbl_d_') == 0, :) = [];

% Write the results table to an Excel file
outputExcelFileName = 'output_results_AvgDlyOil_bbl_d__horizontals_sum.xlsx';  % Replace with your desired output file name
writetable(resultsTable, outputExcelFileName);