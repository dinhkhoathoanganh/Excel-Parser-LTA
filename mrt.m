function mrt
% fpath='D:/Documents/Documents/MRT.xlsx';
% loadData(fpath);
in = load('MRT.mat');
out = struct();
nameList = getNames();
for count = 1:size(nameList,1)
    out(count).data = cell(40,3);
    out(count).data{1,1} = 'Roles';
    out(count).data{1,2} = 'Years';
    out(count).data{1,3} = 'Boards';
end

matrixList = [1 4 6]; % sheets in matrix style
for count = matrixList
    readMatrix(nameList, count);
end

tableList = [2 3 5 7]; % sheets in table style
for count = tableList
    readTable(nameList, count);
end

for count = 1:size(nameList,1) % delete redundant rows
    out(count).data(all(cellfun('isempty',out(count).data),2),:) = [];
end

xlswrite('outMRT.xlsx',nameList,'Names');

for count=1:numel(out)
    sheetName = cell2mat(nameList(count));
    sheetName(regexp(sheetName,'[:,/,?,*]'))=[];
    sheetName(regexp(sheetName,'\s+'))=[];
    xlswrite('outMRT.xlsx',out(count).data,...
        [sheetName(1:6) '(' num2str(size(out(count).data,1)-1) ')']);
    pause(0.2);
end

    function loadData(fpath)
        [~,sheets] = xlsfinfo(fpath);
        for i = 1:size(sheets,2)
            [~,~,data(i).raw] = xlsread(fpath,sheets{i});
        end
        save('MRT.mat');
    end
    function vipNames = getNames
        %% 1st sheet
        temp = in.data(1).raw;
        names={};
        for j=2:size(temp,2)
            for i=2:size(temp,1)-4
                rawN = cell2mat(temp(i,j));
                if ~isnan(rawN)
                    if ~isempty(strfind(rawN,'('))
                        idx = strfind(rawN,'(');
                        rawN = rawN(1:idx(1)-1);
                    end
                    names=[strtrim(rawN); names];
                end
            end
        end
        
        %% 2nd sheet
        temp = in.data(2).raw(:,1);
        for i=2:size(temp,1)
            if ~isnan(cell2mat(temp(i,1)))
                names=[strtrim(temp(i,1)); names];
            end
        end
        
        %% 3rd sheet
        temp = in.data(3).raw(:,1);
        for i=2:31
            rawN = cell2mat(temp(i,1));
            if ~isnan(rawN)
                if strcmp(rawN(1:3),'Dr ')||strcmp(rawN(1:3),'Mr ')
                    rawN = rawN(4:end);
                end
                names=[strtrim(rawN); names];
            end
        end
        
        %% 4th sheet
        temp = in.data(4).raw;
        for j=2:size(temp,2)
            for i=2:size(temp,1)-9
                rawN = cell2mat(temp(i,j));
                if ~isnan(rawN)
                    lengthN = length(rawN);
                    if lengthN>11 && strcmp(rawN(1:10),'Professor ')
                        rawN = rawN(11:end);
                    elseif lengthN>3 && strcmp(rawN(1:3),'Dr ')
                        rawN = rawN(4:end);
                    elseif lengthN>21 && strcmp(rawN(1:20),'Associate Professor ')
                        rawN = rawN(21:end);
                    end
                    names=[strtrim(rawN); names];
                end
            end
        end
        
        %% 5th sheet
        temp = in.data(5).raw(:,1);
        for i=2:size(temp,1)
            rawN = cell2mat(temp(i,1));
            if ~isnan(rawN)
                if strcmp(rawN(1:3),'Dr ')
                    rawN = rawN(4:end);
                end
                names=[strtrim(rawN); names];
            end
        end
        
        %% 6th sheet
        temp = in.data(6).raw;
        for j=2:size(temp,2)
            for i=2:size(temp,1)-2
                rawN = cell2mat(temp(i,j));
                if ~isnan(rawN)
                    lengthN = length(rawN);
                    if strcmp(rawN(1:3),'Dr ')||strcmp(rawN(1:3),'Mr ')||strcmp(rawN(1:3),'Ms ')
                        rawN = rawN(4:end);
                    elseif lengthN>11 && (strcmp(rawN(1:10),'Professor ')||strcmpi(rawN(1:10),'Col (Ret) ')||strcmpi(rawN(1:10),'Radm (NS) '))
                        rawN = rawN(11:end);
                    elseif strcmp(rawN(1:4),'Mrs ')||strcmp(rawN(1:4),'Mdm ')||strcmp(rawN(1:4),'Col ')
                        rawN = rawN(5:end);
                    elseif lengthN>21 && strcmp(rawN(1:20),'Associate Professor ')
                        rawN = rawN(21:end);
                    end
                    if ~isempty(strfind(rawN,'-'))
                        idx = strfind(rawN,'-');
                        rawN = rawN(1:idx(1)-1);
                    end
                    if ~isempty(strfind(rawN,'('))
                        idx = strfind(rawN,'(');
                        rawN = rawN(1:idx(1)-1);
                    end
                    names=[strtrim(rawN); names];
                end
            end
        end
        
        %% 7th sheet
        temp = in.data(7).raw(:,1);
        for i=2:size(temp,1)
            rawN = cell2mat(temp(i,1));
            if ~isnan(rawN)
                if strcmp(rawN(1:3),'Mr ')
                    rawN = rawN(4:end);
                end
                names=[strtrim(rawN); names];
            end
        end
        
        [uniqueNames, ~, org] = unique(names);
        occurences = hist(org,length(uniqueNames));
        duplicateNames = find(occurences>1);
        
        % Processing name lists
        vipNames = [];
        for i=1:size(duplicateNames,2)
            temp = uniqueNames(duplicateNames(i));
            if length(temp{1})>6 % remove source and names headers
                vipNames = [temp; vipNames];
            end
        end
    end
    function readMatrix(vipNames, sheetNo)
        temp = in.data(sheetNo).raw;
        for k = 1:size(vipNames,1)
            searchName = cell2mat(vipNames(k));
            for j=2:size(temp,2)
                for i=2:size(temp,1)
                    rawN = cell2mat(temp(i,j));
                    if ~isempty(strfind(rawN,searchName))
                        idx = find(cellfun('isempty', out(k).data),1);
                        out(k).data{idx,1} = cell2mat(temp(i,1)); %position
                        out(k).data{idx,2} = cell2mat(temp(1,j)); %year
                        out(k).data{idx,3} = in.sheets{sheetNo};
                    end
                end
            end
        end
    end
    function readTable(vipNames, sheetNo)
        temp = in.data(sheetNo).raw;
        for k = 1:size(vipNames,1)
            searchName = cell2mat(vipNames(k));
            for i=2:size(temp,1)
                rawN = cell2mat(temp(i,1));
                if ~isempty(strfind(rawN,searchName))
                    idx = find(cellfun('isempty', out(k).data),1);
                    out(k).data{idx,1} = cell2mat(temp(i,2)); %position
                    out(k).data{idx,2} = cell2mat(temp(i,3)); %year
                    out(k).data{idx,3} = in.sheets{sheetNo};
                end
            end
        end
    end
end
        
