# Stock-Summarizer
Stock Summarizer using VBA

# User Instructions:

To use the workbook for generating stock summary; 2 options - 
    1. Through the worksheet “Control Sheet” – First worksheet of the workbook. 
        -	Button “Summarize the stocks” – on click of this button, all the sheets in the workbook would be summarized into corresponding worksheet.
        -	Button “Reset to Originals” – on click of this button, the new column details added to each stock sheet as part of summary would be cleared.
    2. Individual stock sheets have 2 buttons to summarize the single active worksheet and/or to clear the worksheet.

To use vbs code, refer below recommendations.

    1.	Create a common worksheet with name “Control Sheet” and add buttons.
        Assign macros “SummarizeAllStockSheets” and “ResetAllSheets” to each button.

    2.	Create 2 buttons in each stock sheet to summarize individual sheet.
        Assign macros “SummarizeStocks” and “ResetCurrentSheet” to each button.

Important subroutines for stock summary
-	SummarizeAllStockSheets – subroutine helps summarize the stock data in all worksheets of the spreadsheet.
-	SummarizeStocks –subroutine helps summarize stock data from only one active worksheet.
-	ResetAllSheets – Clears all the columns that got added as part of overall summary in all the worksheets of the workbook.
-	ResetCurrentSheet – clears additional columns added as part of SummarizeStocks in the current sheet.

# Assumptions:
1.	All stock details associated with a Ticker are grouped together into single worksheet.
2.	Each worksheet has data for only one financial year.
3.	Ticker details are sorted by the date in ascending order.
