# AutomaticRefreshExcelFile
This project is to solve, how to refresh the Excel file which is stored on SharePoint

All our company's files are stored on Sharepoint. Some of these Excel files are linked to external data and therefore cannot be refreshed directly on Sharepoint. So each time we need to open excel in the app and click RefreshAll then save file. We have many files that need to be refreshed on a daily basis like this. To solve the problem of repetition, I use the following method.

1. Sync SharePoint files to Local via OneDrive
2. Use Python or Powershell to solve the issue
3. Add the program to Task Scheduler

For PowerShell solution, you can reference the orginal link https://github.com/TylerNielsen/powershell-refresh-excel/blob/master/RefreshExcelFiles.ps1
In my code I have added a time stamp to detect when a file has failure.

It is worth noting that you must ensure that your 'Enable background refresh' function is unchecked.

![grafik](https://user-images.githubusercontent.com/84840321/159511306-9dbd461c-daaa-4ee3-8c1b-c896e84b20e6.png)

Otherwise you will get the following error

![grafik](https://user-images.githubusercontent.com/84840321/159511471-604f61ad-0817-4c79-9efe-79ce46342d53.png)


