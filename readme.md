# Excel Tools

This repository contains macros and other macro-related code clips for use in Excel.

## TreeNode.cls

This class represents a nestable tree node object. 


## [Automatic Backup.txt](https://github.com/dondowjr/ExcelTools-repo/blob/main/Automatic%20Backup.txt)

This block of code enables automatically backing up an Excel workbook when it is first opened.

For workbooks on a local driver or a network share, a copy of the workbook is made and placed into a 'backups' directory below the location of the opened workbook.

For workbooks from an online source, such as SharePoint or OneDrive, a copy is made and placed into the user's 'Documents' library under a folder called 'Office Backups'.
The file name of the copy is the same as the original except that it has a date stamp appended to the end.

When a workbook with this code is opened, it will check if a backup has been made for the day. If not, the user will be prompted to back up the file.  

If a backup is made, the code checks if there are old copies that can be purged. If old copies are found, the user is prompted to purge the old backups.

Inside the code, there is a custom settings block that allows you to specify how many copies to keep and whether to always back up to the 'Documents' library.

To implement this code, place the code bock into the code window for the ‘ThisWorkbook’ Class.


