% This section of script gives last value and sum of all values for each
% well for a given column
% Replace 'your_file.xlsx' with the actual name of your Excel file
excelFileName = 'Production_History.xlsx';

% Read the Excel file into a table
dataTable = readtable(excelFileName);

% Extract the column with the title 'Sort UWI'
sortUWIColumn = dataTable.('SortUWI');

% Remove empty or 0x0 character values
sortUWIColumn = sortUWIColumn(~cellfun('isempty', sortUWIColumn));

% Count unique values in the 'Sort UWI' column
uniqueValues = unique(sortUWIColumn);

% Create a cell array to store the results
resultsCellArray = {'SortUWI', 'Case A Sum', 'Case B Last Value'};

% Loop through each unique value in 'Sort UWI'
for i = 1:length(uniqueValues)
    currentSortUWI = uniqueValues{i};
    
    % Case A: Get the sum of all values in a specific column
    specificColumn = 'CumPrdWater_bbl_';
    idxA = strcmp(dataTable.SortUWI, currentSortUWI);
    sumValueA = nansum(dataTable.(specificColumn)(idxA)); % Use nansum to handle NaNs
    
    disp(['Case A - SortUWI: ', currentSortUWI, ', Sum of ', specificColumn, ': ', num2str(sumValueA)]);
    
    % Case B: Get the last value in a column for a specific time
    specificTime = datetime('31-Dec-1975', 'Format', 'dd-MMM-yyyy');
    idxB = strcmp(dataTable.SortUWI, currentSortUWI) & (dataTable.ProdDate <= specificTime);
    lastValueB = dataTable.(specificColumn)(idxB);
    
    % Check if lastValueB is not empty before attempting to access its last element
    if ~isempty(lastValueB)
        lastValueB = lastValueB(end);
        disp(['Case B - SortUWI: ', currentSortUWI, ', Last value of ', specificColumn, ' before ', datestr(specificTime), ': ', num2str(lastValueB)]);
    else
        disp(['Case B - SortUWI: ', currentSortUWI, ', No values of ', specificColumn, ' before ', datestr(specificTime)]);
        lastValueB = 0;  % Change the last value to zero if it is empty
    end
    
    % Store results in the cell array
    resultsCellArray = [resultsCellArray; {currentSortUWI, sumValueA, lastValueB}];
end

% ... (Your existing code)

% Convert the cell array to a table
resultsTable = cell2table(resultsCellArray(2:end, :), 'VariableNames', resultsCellArray(1, :));

% Write the results table to an Excel file
outputExcelFileName = 'output_results_vprd.xlsx';  % Replace with your desired output file name
writetable(resultsTable, outputExcelFileName);

% % Loop through each unique value in 'Sort UWI'
% for i = 1:length(uniqueValues)
%     currentSortUWI = uniqueValues{i};
%     
%     % Case A: Get the sum of all values in a specific column
%     specificColumn = 'CumPrdOil_bbl_';
%     idxA = strcmp(dataTable.SortUWI, currentSortUWI);
%     sumValueA = nansum(dataTable.(specificColumn)(idxA)); % Use nansum to handle NaNs
%     
%     disp(['Case A - SortUWI: ', currentSortUWI, ', Sum of ', specificColumn, ': ', num2str(sumValueA)]);
%     
%     % Case B: Get the last value in a column for a specific time
%     specificTime = datetime('01-Jan-1975', 'Format', 'dd-MMM-yyyy');
%     idxB = strcmp(dataTable.SortUWI, currentSortUWI) & (dataTable.ProdDate <= specificTime);
%     lastValueB = dataTable.(specificColumn)(idxB);
%     
%     % Check if lastValueB is not empty before attempting to access its last element
%     if ~isempty(lastValueB)
%         lastValueB = lastValueB(end);
%         disp(['Case B - SortUWI: ', currentSortUWI, ', Last value of ', specificColumn, ' before ', datestr(specificTime), ': ', num2str(lastValueB)]);
%     else
%         disp(['Case B - SortUWI: ', currentSortUWI, ', No values of ', specificColumn, ' before ', datestr(specificTime)]);
%         %lastValueB = 0 ;
%     end
% end
% 
% 
% % Convert the cell array to a table
% resultsTable = cell2table(resultsCellArray(2:end, :), 'VariableNames', resultsCellArray(1, :));
% 
% % Write the results table to an Excel file
% outputExcelFileName = 'output_results_vprd.xlsx';  % Replace with your desired output file name
% writetable(resultsTable, outputExcelFileName);

% % Loop through each unique value in 'Sort UWI'
% for i = 1:length(uniqueValues)
%     currentSortUWI = uniqueValues{i};
%     
%     % Case A: Get the sum of all values in a specific column
%     specificColumn = 'CumPrdOil_bbl_';
%     idxA = strcmp(dataTable.SortUWI, currentSortUWI);
%     sumValueA = nansum(dataTable.(specificColumn)(idxA));
%     
%     disp(['Case A - SortUWI: ', currentSortUWI, ', Sum of ', specificColumn, ': ', num2str(sumValueA)]);
%     
%     % Case B: Get the last value in a column for a specific time
%     specificTime = datetime('01-Jan-1975', 'Format', 'dd-MMM-yyyy');
%     idxB = strcmp(dataTable.SortUWI, currentSortUWI) & (dataTable.ProdDate <= specificTime);
%     lastValueB = dataTable.(specificColumn)(idxB);
%     lastValueB = lastValueB(end);
%     
%     disp(['Case B - SortUWI: ', currentSortUWI, ', Last value of ', specificColumn, ' before ', datestr(specificTime), ': ', num2str(lastValueB)]);
% end

% ... (Your existing code)