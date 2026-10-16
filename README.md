# Remittance Report Extractor

Run `python main.py` from the project directory to process every `.xls`, `.xlsx`,
and `.xlsm` workbook in the configured input folder. For each input workbook,
the bot copies `Output_Template/OutputFile_Temp2.xlsx`, fills the copied template,
and saves it under `output/` using the input workbook name followed by
`Output File.xlsx`. A failure for one input is logged while processing continues
for the remaining workbooks; the run exits with an error status if any workbook
could not be processed.

Each generated workbook retains the template's worksheets and formatting, with
empty `Payments`, `Cost`, or `IntOnly` worksheets removed when they have no
output rows. Summary formulas that reference a removed worksheet are set to
zero. The `Summary1` direct-payment commission is populated from the remittance
statement. Payment rows with non-zero allocated totals are included, except
interest-only payments; a missing reference number is left blank in the output.
The payment total is principal, interest, fees, costs, and other amounts less
direct payments. Applied principal is principal plus fees less direct payments.
The output's `CLIENT SHARE` column
uses the master-table `BWShare` value, while `BW, LLP SHARE` uses `ClientShare`,
matching the expected payment workbook. Interest uses its BW interest percentage.

The `Remittance Report` payment section is processed before the `Cost Detail
Report` section. These sections are handled independently: if either section
or its data rows are missing, that section is skipped with a warning and the
other section is still written to the output workbook.

The client share, BW LLP share, and BW interest percentage are looked up for each
input filename in `MasterDataTable/MasterDatatable_1.xlsx`, sheet
`IntrestPercentage_Temp2`. Missing or duplicate client entries stop that file
with a logged error; percentages are not silently defaulted.

Input/output paths, the template, and the master workbook are configured in
`config/config.xlsx`.