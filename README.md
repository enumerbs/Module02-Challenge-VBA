# Module02-Challenge-VBA
Data Analytics Boot Camp - Module 02 - VBA Scripting \
VBA Challenge

---

# Results

Screenshots of the results for each worksheet (per-Quarter stock summary and 'greatest' statistics) are given in the following files in the **screenshots** subfolder:
- screenshot_VBA_Challenge_Result_Worksheet_Q1.png
- screenshot_VBA_Challenge_Result_Worksheet_Q2.png
- screenshot_VBA_Challenge_Result_Worksheet_Q3.png
- screenshot_VBA_Challenge_Result_Worksheet_Q4.png

# Implementation notes

The '.xlsx' files provided in this Challenge's Starter_Code were converted to '.xlsm' so they could store VBA source code.

Implentation development and testing were performed using the supplied *alphabetical_testing* data set. The '.xlsm' file and the extracted VBA code '.vbs' file from that stage are archived in the 'alphabetical_testing' subfolder.

Additional code comments were added to that VBA code to produce the final version saved as **CalculateAndDisplayStockInformation.bas**. That VBA code was loaded into Module1 in **Multiple_year_stock_data.xlsm** and run from there to produce the final results.

### *Assumption*

In order to match the sample result provided in the 'Bonus' section of this Challenge, the 'greatest' stock is calculated on a per-worksheet (per-quarter) basis, not the 'greatest' overall.



# References

The following references were used in the development of the solution for this Challenge.

## Markdown
*Used in developing the contents of this README.md file!*
- https://www.markdownguide.org/cheat-sheet/
- https://code.visualstudio.com/docs/languages/markdown#_markdown-preview

## Git & GitHub
- **git mv** https://git-scm.com/docs/git-mv
- **Adding your SSH key to the ssh-agent** https://docs.github.com/en/authentication/connecting-to-github-with-ssh/generating-a-new-ssh-key-and-adding-it-to-the-ssh-agent#adding-your-ssh-key-to-the-ssh-agent

## Excel VBA

### Import .vbs file into Excel VBA editor
- https://stackoverflow.com/questions/56680417/how-to-import-a-macro-to-excel-by-vbs
- https://www.reddit.com/r/vba/comments/1c5g8s/how_to_import_and_run_external_macro_bas_file/
- https://support.microsoft.com/en-au/office/work-with-vba-macros-in-excel-for-the-web-98784ad0-898c-43aa-a1da-4f0fb5014343

### Export macro to .vbs file
- https://stackoverflow.com/questions/43352194/how-to-save-macro-code-as-vbs
- https://www.mrexcel.com/archive/vba/saving-an-excel-macro-as-a-vbs-file/

### Excel VBA get used range
- https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.usedrange
- https://www.wallstreetmojo.com/vba-usedrange/

### Excel VBA get worksheets
- https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet#properties
- https://trumpexcel.com/vba-worksheets/
- https://excelmacromastery.com/excel-vba-worksheet/

### Excel VBA call sub with parameters
- https://stackoverflow.com/questions/56259496/calling-a-sub-in-vba-with-multiple-arguments
- https://www.wallstreetmojo.com/vba-call-sub/

### Excel VBA module global constant
- https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/declaring-constants
- https://www.mrexcel.com/board/threads/global-const-vba-best-practices.362948/

### Excel VBA cell number format
- https://learn.microsoft.com/en-us/office/vba/api/excel.cellformat.numberformat

### Excel VBA cell percentage format
- https://www.statology.org/vba-percentage-format/

### Excel VBA cell background color
- https://stackoverflow.com/questions/365125/how-do-i-set-the-background-color-of-excel-cells-using-vba
- https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)?redirectedfrom=MSDN
- https://www.excel-easy.com/vba/examples/background-colors.html

### Excel VBA Autofit
- https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit
- https://excelchamps.com/vba/autofit/