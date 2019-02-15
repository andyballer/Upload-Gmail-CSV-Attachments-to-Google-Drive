# Upload-Gmail-CSV-Attachments-to-Google-Drive
Uses Google Apps Scripts to automate uploading gmail attachments to Google Drive

Run runEveryDay in Google Apps Scripts to trigger uploadSpreadsheetToDrive every day between 1am and 2am, replacing your old data with 
current data. Running uploadDataToDrive gets a specified number of spreadsheets from your emails, and uploads them to your specified 
root Drive folder based on the search returned. Once the data is in Google Drive, you can access visualize your data with tools like 
Splunk and Google Data Studio. 

*Need to set the variables before running.* 
