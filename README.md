# XLSWRITEFIG

Write a MATLAB figure to an Excel spreadsheet

## Syntax
```
xlswritefig(hFig, filename, sheetname, xlcell)
```

## Input Arguments

All inputs are optional:

| Name | Description |
| --- | --- |
| hFig | Handle to MATLAB figure.  If empty, current figure is exported |
| filename | (string) Name of Excel file, including extension.  If not specified, contents will be opened in a new Excel spreadsheet. |
| sheetname |  Name of sheet to write data to. The default is 'Sheet1'. If specified, a sheet with the specified name must exist. |
| xlcell | Designation of cell to indicate the upper-left corner of the figure (e.g. 'D2').  Default = 'A1' |

## Requirements
- Must have Microsoft Excel installed.
- Microsoft Windows only.

## Examples
 
Paste the current figure into a new Excel spreadsheet which is left open.
``` 
plot(rand(10,1))
xlswritefig
```

Specify all options.  
```
hFig = figure;      
surf(peaks)
xlswritefig(hFig,'MyNewFile.xlsx','Sheet2','D4')
winopen('MyNewFile.xlsx') 
```
