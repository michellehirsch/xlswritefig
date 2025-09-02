function xlswritefig(hFig,filename,sheetname,xlcell)
% XLSWRITEFIG  Write a MATLAB figure to an Excel spreadsheet
%
% xlswritefig(hFig,filename,sheetname,xlcell)
%
% All inputs are optional:
%
%    hFig:      Handle to MATLAB figure.  If empty, current figure is
%                   exported
%    filename   (string) Name of Excel file, including extension.  If not specified, contents will
%                  be opened in a new Excel spreadsheet. 
%    sheetname:  Name of sheet to write data to. The default is 'Sheet1'
%                       If specified, a sheet with the specified name must
%                       exist
%    xlcell:     Designation of cell to indicate the upper-left corner of
%                  the figure (e.g. 'D2').  Default = 'A1'
%
% Requirements: Must have Microsoft Excel installed.  Microsoft Windows
% only.
%
% Ex:
% Paste the current figure into a new Excel spreadsheet which is left open.
%         plot(rand(10,1))
%         drawnow    % Maybe overkill, but ensures plot is drawn first
%         xlswritefig
%
% Specify all options.  
%         hFig = figure;      
%         surf(peaks)
%         xlswritefig(hFig,'MyNewFile.xlsx','Sheet2','D4')
%         winopen('MyNewFile.xlsx')   

% Michelle Hirsch
% The MathWorks
% mhirsch@mathworks.com
%
% Is this function useful?  Drop me a line to let me know!


if nargin==0 || isempty(hFig)
    hFig = gcf;
end

if nargin<2 || isempty(filename)
    filename ='';
    dontsave = true;
else
    dontsave = false;
    
    % Create full file name with path
    filename = fullfilename(filename);
end

if nargin < 3 || isempty(sheetname)
    sheetname = 'Sheet1';
end

if nargin<4
    xlcell = 'A1';
end


% Put figure in clipboard
if ~verLessThan('matlab','9.8')
    warning off MATLAB:print:ExportExcludesUI
    copygraphics(hFig)
    warning on MATLAB:print:ExportExcludesUI
else
    % For older releases, use hgexport. Set renderer to painters to make
    % sure it looks right.
    r = get(hFig,'Renderer');
    set(hFig,'Renderer','Painters')
    drawnow
    hgexport(hFig,'-clipboard') %#ok<HGEXPORT>
    set(hFig,'Renderer',r)
end


% Open Excel, add workbook, change active worksheet,
% get/put array, save.
% First, open an Excel Server.
Excel = actxserver('Excel.Application');

% Two cases:
% * Open a new workbook, save with given file name
% * Open an existing workbook

if exist(filename,'file')==0
    % The following case if file does not exist (Creating New File)
    op = invoke(Excel.Workbooks,'Add');
    %     invoke(op, 'SaveAs', [pwd filesep filename]);
    new=1;
else
    % The following case if file does exist (Opening File)
    %     disp(['Opening Excel File ...(' filename ')']);
    op = invoke(Excel.Workbooks, 'open', filename);
    new=0;
end

% set(Excel, 'Visible', 0);

% Make the specified sheet active.
try
    Sheets = Excel.ActiveWorkBook.Sheets;
    target_sheet = get(Sheets, 'Item', sheetname);
catch %#ok<CTCH>   Suppress so that this function works in releases without MException
    % Add the sheet if it doesn't exist
    target_sheet = Excel.ActiveWorkBook.Worksheets.Add();
    target_sheet.Name = sheetname;

end

invoke(target_sheet, 'Activate');
Activesheet = Excel.Activesheet;


% --------------------
% Try clipboard paste first; on failure, insert from a file (robust)
% --------------------
try
    % Paste to specified cell
    Paste(Activesheet,get(Activesheet,'Range',xlcell,xlcell));
catch %#ok<CTCH>
    % Fallback: export to file and insert without clipboard
    
    % USe EMF vector on Windows for crisp lines/text
    % Use .png if images - I haven't added an input option, so you are on your own
    tmpfile = [tempname '.emf'];
    if ~verLessThan('matlab','9.8')
        warning("off", 'MATLAB:print:ContentTypelmageSuggested')
        exportgraphics(hFig,tmpfile,'ContentType','vector');
        warning("on", 'MATLAB:print:ContentTypelmageSuggested')
    else
        print(hFig,tmpfile,'-dmeta'); % EMF via print
    end

    % Insert and position at the anchor cell
    anchor = get(Activesheet,'Range',xlcell,xlcell);
    left = anchor.Left;  top = anchor.Top;
    Shapes = Activesheet.Shapes;
    % LinkToFile=0 (msoFalse), SaveWithDocument=1 (msoTrue), Width/Height=-1 keep native size
    pic = invoke(Shapes,'AddPicture', tmpfile, 0, 1, left, top, -1, -1);
    try %#ok<TRYNC>
        % Lock aspect ratio when available
        pic.LockAspectRatio = true;
    end

    if ~isempty(tmpfile) && exist(tmpfile,'file'), delete(tmpfile); end
end



% Save and clean up
if new && ~dontsave
    invoke(op, 'SaveAs', filename);
elseif ~new
    invoke(op, 'Save');
else  % New, but don't save
    set(Excel, 'Visible', 1);
    return  % Bail out before quitting Excel
end
invoke(Excel, 'Quit');
delete(Excel)
end


function filename = fullfilename(filename)
[filepath, filename, fileext] = fileparts(filename);
if isempty(filepath)
    filepath = pwd;
end
if isempty(fileext)
    fileext = '.xlsx';
end
filename = fullfile(filepath, [filename fileext]);
end
