using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;

namespace ExcelHandler
{
    /// <summary>
    /// This class updates external data connections in excel and update data in excel file
    /// </summary>
    public static class ExternalDataUpdater
    {
        public static void UpdateSharepointFile(string excelSharepointFilePath)
        {
            Excel.Application excelApp = null;
            string excelLocalWorkBookName = null;
            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(excelSharepointFilePath,
                    0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);

                excelWorkbook.RefreshAll();
                System.Threading.Thread.Sleep(5000);

                excelLocalWorkBookName = "temp_" + excelWorkbook.Name;

                string path = Path.Combine(Directory.GetCurrentDirectory(), excelLocalWorkBookName);
                excelWorkbook.SaveAs(path, ConflictResolution: Excel.XlSaveConflictResolution.xlLocalSessionChanges);


                excelWorkbook.Close(true);
                PublishWorkbook(path, excelSharepointFilePath);
                //excelApp.CalculateUntilAsyncQueriesDone();


            }
            finally
            {
                if (excelApp != null)
                {
                    excelApp.Quit();
                    excelApp = null;
                }
                if (excelLocalWorkBookName != null)
                {
                    if (File.Exists(excelLocalWorkBookName))
                    {
                        File.Delete(excelLocalWorkBookName);
                    }
                    else
                    {
                        throw new Exception("Weird! We have temp file, but can not delete it, because it was not found");
                    }
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        static void PublishWorkbook(string LocalPath, string SharePointPath)
        {
            WebResponse response = null;

            try
            {
                // Create a PUT Web request to upload the file.
                WebRequest request = WebRequest.Create(SharePointPath);

                request.Credentials = CredentialCache.DefaultCredentials;
                request.Method = "PUT";

                // Allocate a 1K buffer to transfer the file contents.
                // The buffer size can be adjusted as needed depending on
                // the number and size of files being uploaded.
                byte[] buffer = new byte[1024];

                // Write the contents of the local file to the
                // request stream.
                using (Stream stream = request.GetRequestStream())
                using (FileStream fsWorkbook = File.Open(LocalPath,
                    FileMode.Open, FileAccess.Read))
                {
                    int i = fsWorkbook.Read(buffer, 0, buffer.Length);

                    while (i > 0)
                    {
                        stream.Write(buffer, 0, i);
                        i = fsWorkbook.Read(buffer, 0, buffer.Length);
                    }
                }

                // Make the PUT request.
                response = request.GetResponse();
            }
            finally
            {
                response.Close();
            }
        }

        static IEnumerable<string> GetSharepointPaths (string libraryUrl)
        {
            throw new NotImplementedException();
            //using (SPSite site = new SPSite("http://localhost/sites/sitecollection"))
            //{

            //}
        }
    }

}
