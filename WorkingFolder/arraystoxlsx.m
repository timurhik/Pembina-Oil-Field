% %This section of script counts number of unique wells
% % Replace 'your_file.xlsx' with the actual name of your Excel file
% %excelFileName = 'Injection_History.xlsx';
% %excelFileName = 'Production_History.xlsx';
% %excelFileName = 'Horizontals_Production_History.xlsx';
% 
% % Read the Excel file into a table
% dataTable = readtable(excelFileName)
% 
% % Extract the column with the title 'Sort UWI'
% sortUWIColumn = dataTable.('SortUWI');
% 
% % Remove empty or 0x0 character values
% sortUWIColumn = sortUWIColumn(~cellfun('isempty', sortUWIColumn));
% 
% % Count unique values in the 'Sort UWI' column
% uniqueValues = unique(sortUWIColumn);
% numUniqueValues = length(uniqueValues);
% 
% % Display the unique values and the count
% disp('Unique Values:');
% disp(uniqueValues);
% disp(['Number of Unique Values: ', num2str(numUniqueValues)]);

%-------------------------------------------------------------------------
% % This section of script gives last value and sum of all values for each
% % well for a given column
% % Replace 'your_file.xlsx' with the actual name of your Excel file
% excelFileName = 'Injection_History.xlsx';
% 
% % Read the Excel file into a table
% dataTable = readtable(excelFileName);
% 
% % Extract the column with the title 'Sort UWI'
% sortUWIColumn = dataTable.('SortUWI');
% 
% % Remove empty or 0x0 character values
% sortUWIColumn = sortUWIColumn(~cellfun('isempty', sortUWIColumn));
% 
% % Count unique values in the 'Sort UWI' column
% uniqueValues = unique(sortUWIColumn);
% 
% % Loop through each unique value in 'Sort UWI'
% for i = 1:length(uniqueValues)
%     currentSortUWI = uniqueValues{i};
%     
%     % Case A: Get the sum of all values in a specific column
%     specificColumn = 'CumInjWater_bbl_';
%     idxA = strcmp(dataTable.SortUWI, currentSortUWI);
%     sumValueA = sum(dataTable.(specificColumn)(idxA));
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
% 
% % ... (Your existing code)
% 
% % Create a cell array to store the results
% resultsCellArray = {'SortUWI', 'Case A Sum', 'Case B Last Value'};
% 
% % Loop through each unique value in 'Sort UWI'
% for i = 1:length(uniqueValues)
%     currentSortUWI = uniqueValues{i};
%     
%     % Case A: Get the sum of all values in a specific column
%     specificColumn = 'CumInjWater_bbl_';
%     idxA = strcmp(dataTable.SortUWI, currentSortUWI);
%     sumValueA = sum(dataTable.(specificColumn)(idxA));
%     
%     % Case B: Get the last value in a column for a specific time
%     specificTime = datetime('01-Jan-1975', 'Format', 'dd-MMM-yyyy');
%     idxB = strcmp(dataTable.SortUWI, currentSortUWI) & (dataTable.ProdDate <= specificTime);
%     lastValueB = dataTable.(specificColumn)(idxB);
%     lastValueB = lastValueB(end);
%     
%     % Display and store the results
%     disp(['SortUWI: ', currentSortUWI, ', Sum of ', specificColumn, ': ', num2str(sumValueA), ...
%           ', Last value of ', specificColumn, ' before ', datestr(specificTime), ': ', num2str(lastValueB)]);
%     
%     % Store results in the cell array
%     resultsCellArray = [resultsCellArray; {currentSortUWI, sumValueA, lastValueB}];
% end
% 
% % Convert the cell array to a table
% resultsTable = cell2table(resultsCellArray(2:end, :), 'VariableNames', resultsCellArray(1, :));
% 
% % Write the results table to an Excel file
% outputExcelFileName = 'output_results.xlsx';  % Replace with your desired output file name
% writetable(resultsTable, outputExcelFileName);

%-------------------------------------------------------------------------
% This section of script gives the time frame of all injector wells in
% which their cumm_inj_water was non-zero

% Replace 'your_file.xlsx' with the actual name of your Excel file
excelFileName = 'Horizontals_Production_History.xlsx'; %'Injection_History.xlsx';

% Read the Excel file into a table
dataTable = readtable(excelFileName);

% Extract the column with the title 'Sort UWI'
sortUWIColumn = dataTable.('SortUWI');

% Remove empty or 0x0 character values
sortUWIColumn = sortUWIColumn(~cellfun('isempty', sortUWIColumn));

% Count unique values in the 'Sort UWI' column
uniqueValues = unique(sortUWIColumn);

% Initialize common time frame variables
commonStartTime = NaT;
commonEndTime = NaT;

% Loop through each unique value in 'Sort UWI'
for i = 1:length(uniqueValues)
    currentSortUWI = uniqueValues{i};
    
    % Find rows for the current SortUWI with non-zero CumInjWater_bbl_ values
    idxNonZero = strcmp(dataTable.SortUWI, currentSortUWI) & (dataTable.CumPrdOil_bbl_ > 0);
    currentSortUWIData = dataTable(idxNonZero, :);
    
    % Find the common time frame
    if ~isempty(currentSortUWIData)
        startTime = min(currentSortUWIData.ProdDate);
        endTime = max(currentSortUWIData.ProdDate);
        
        % Update common start and end times
        if isempty(commonStartTime) || startTime > commonStartTime
            commonStartTime = startTime;
        end
        
        if isempty(commonEndTime) || endTime < commonEndTime
            commonEndTime = endTime;
        end
    end
    
    % Display active time for each unique SortUWI value
    disp(['SortUWI: ', currentSortUWI, ', Active Time - Start Time: ', datestr(startTime, 'dd-mmm-yyyy'), ', End Time: ', datestr(endTime, 'dd-mmm-yyyy')]);
end

% Display the common time frame
disp(['Common Time Frame - Start Time: ', datestr(commonStartTime, 'dd-mmm-yyyy'), ', End Time: ', datestr(commonEndTime, 'dd-mmm-yyyy')]);

%-------------------------------------------------------------------------

% % % This section of script gives the activity of all wells based on monthly
% % water injection for each well. 
% % Replace 'your_file.xlsx' with the actual name of your Excel file
% excelFileName = 'Injection_History.xlsx';
% 
% % Read the Excel file into a table
% dataTable = readtable(excelFileName);
% 
% % Extract the column with the title 'Sort UWI'
% sortUWIColumn = dataTable.('SortUWI');
% 
% % Remove empty or 0x0 character values
% sortUWIColumn = sortUWIColumn(~cellfun('isempty', sortUWIColumn));
% 
% % Count unique values in the 'Sort UWI' column
% uniqueValues = unique(sortUWIColumn);
% 
% % Initialize group time frame variables
% groupStartTime = NaT;
% groupEndTime = NaT;
% 
% % Loop through each unique value in 'Sort UWI'
% for i = 1:length(uniqueValues)
%     currentSortUWI = uniqueValues{i};
%     
%     % Find rows for the current SortUWI with non-zero MonInjWater_bbl_ values
%     idxNonZero = strcmp(dataTable.SortUWI, currentSortUWI) & (dataTable.MonInjWater_bbl_ > 0);
%     currentSortUWIData = dataTable(idxNonZero, :);
%     
%     % Find the active time based on MonInjWater_bbl_ values
%     if ~isempty(currentSortUWIData)
%         startTime = min(currentSortUWIData.ProdDate);
%         endTime = max(currentSortUWIData.ProdDate);
%         
%         % Display active time for each unique SortUWI value
%         disp(['SortUWI: ', currentSortUWI, ', Active Time - Start Time: ', datestr(startTime, 'dd-mmm-yyyy'), ', End Time: ', datestr(endTime, 'dd-mmm-yyyy')]);
%         
%         % Update group start and end times
%         if isempty(groupStartTime) || startTime < groupStartTime
%             groupStartTime = startTime;
%         end
%         
%         if isempty(groupEndTime) || endTime > groupEndTime
%             groupEndTime = endTime;
%         end
%     end
% end
% 
% % This sub-section is not working currently
% % % Display the common time frame for all wells
% % if ~isnat(groupStartTime) && ~isnat(groupEndTime)
% %     disp(['Common Time Frame for All Wells - Group Start Time: ', datestr(groupStartTime, 'dd-mmm-yyyy'), ', Group End Time: ', datestr(groupEndTime, 'dd-mmm-yyyy')]);
% % else
% %     disp('No common time frame found for all wells.');
% % end

%-------------------------------------------------------------------------

% % ... (previous code)
% 
% % Define time periods
% periods = {
%     'Jan 1960 - Dec 1975', 'Jan 1976 - Dec 1990', 'Jan 1991 - Dec 2005', 'Jan 2006 - Dec 2023'
% };
% 
% % Create tables to store the results for each period
% lastValuesTable = table('Size', [length(uniqueValues), 1], 'VariableTypes', {'cell'}, 'VariableNames', {'SortUWI'});
% averageValuesTable = table('Size', [length(uniqueValues), 1], 'VariableTypes', {'cell'}, 'VariableNames', {'SortUWI'});
% 
% % Loop through each unique value in 'Sort UWI'
% for i = 1:length(uniqueValues)
%     currentSortUWI = uniqueValues{i};
%     
%     % Find rows for the current SortUWI with non-zero MonInjWater_bbl_ values
%     idxNonZero = strcmp(dataTable.SortUWI, currentSortUWI) & (dataTable.MonInjWater_bbl_ > 0);
%     currentSortUWIData = dataTable(idxNonZero, :);
%     
%     % Initialize tables for last and average values
%     lastValues = table();
%     averageValues = table();
%     
%     % Loop through each time period
%     for p = 1:length(periods)
%         % Extract start and end dates for the current period
%         periodDates = split(periods{p}, ' - ');
%         startDate = datetime(periodDates{1}, 'InputFormat', 'MMM yyyy');
%         endDate = datetime(periodDates{2}, 'InputFormat', 'MMM yyyy') + calmonths(11);
%         
%         % Filter data for the current time period
%         idxPeriod = currentSortUWIData.ProdDate >= startDate & currentSortUWIData.ProdDate <= endDate;
%         periodData = currentSortUWIData(idxPeriod, :);
%         
%         % Calculate last values for specified columns
%         lastValues.(periods{p}) = table(last(periodData.CumInjWater_bbl_), last(periodData.CumPrdOil_bbl_), last(periodData.CumPrdGas_mcf_));
%         
%         % Calculate average values for specified columns
%         averageValues.(periods{p}) = table(mean(periodData.AvgInjWater_bbl_d_), mean(periodData.AvgDlyOil_bbl_d_), mean(periodData.AvgDlyGas_mcf_d_));
%     end
%     
%     % Combine last and average values tables with SortUWI column
%     lastValuesTable{i, :} = [currentSortUWI, lastValues];
%     averageValuesTable{i, :} = [currentSortUWI, averageValues];
% end
% 
% % Display results for last values
% disp('Last Values for Each SortUWI and Time Period:');
% disp(lastValuesTable);
% 
% % Display results for average values
% disp('Average Values for Each SortUWI and Time Period:');
% disp(averageValuesTable);



