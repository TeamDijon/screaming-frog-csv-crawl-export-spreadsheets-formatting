# screaming-frog-crawl-export-spreadsheets-formatting

A spreadsheet coupled with an Google Apps Script in order to automate Screaming Frog crawl csv export formatting

## How to use the script ?

In order to use the script, just go to the following spreadsheet and make your own copy:

https://docs.google.com/spreadsheets/d/1lyJLZsAILlhmAwWbHAq1ygWnLYOtqUIqnlM03bmJoGE/edit?usp=sharing

> I usually go with the name "ClientName - Crawls" in the appropriate folder

Once you have your file copy, you will need to paste your csv crawl export in the A1 cell of the "First crawl" tab

After that, you will be able to see a "Process" tab in the taskbar of Spreadsheets, the script created the menu for you to click on the "Format Crawl Export" option.

The first time, the script will ask permission to modify the spreadsheets (no worry, no additional permission is granted, no malicious code is inside)

Once you're done, you can click "Format Crawl Export" again and the script will be formatted in a matter of seconds

Now enjoy the crawl audit without any setup necessary


## Fast formatting features

- Split crawl contents into columns
- Apply hyperlinks in place of text for the URLs columns
- Freeze the first row / first column
- Optimize the sheet text formatting
- Change column widths according to the content
- Add a filter on the crawl data
- Change the tab name to the current date
- Add conditional formatting rules to 
- Set the active range on "Content Type" column to prepare for HTML content filtering


## Other things to keep in mind

- The script is optimized for an english csv export, you will have to configure the script accordingly
- Script is commented, making it easier to tweak the different features
- If you have other ideas or questions, do not hesitate !
