# ExcelUpdater
Small util for updating external data connections for excel files located at sharepoint site
Features:
  - Logging
  - Command line options
     - SiteUrl (Required - site where excel files located)
     - LibraryName (Required - library where excel files located)
     - FolderName (Folder in library where excel files located)
     - ExcelVisible (false by default - Show excel instance during updating of files)

TODO:
Now authentication to site is using default credentials. 
Before one starts updating, one should open excel and open Sharepoint library "by hands". 
Excel remembers credentials and will use it in future 

Workflow:
  1. Open local instance of Excel
  2. Open workbook from sharepoint
  3. Save it locally in temp file (otherwise file can not be changed, because it is in read only mode)
  4. Refresh all connections in file
  5. Save it again
  6. Publish back to Sharepoint site
  7. Delete file
  8. Go to 2. if there are other files to update. Else 9
  9. Close Instance of Excel
