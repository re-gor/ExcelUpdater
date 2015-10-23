using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using Microsoft.SharePoint.Client;
using log4net;

namespace ExcelHandler
{
    /// <summary>
    /// This class updates external data connections in excel and update data in excel file
    /// </summary>
    public static class ExternalDataUpdater
    {
        private static readonly ILog _log = LogManager.GetLogger(typeof(ExternalDataUpdater));
        public static void UpdateSharepointFiles(string siteUrl, string libraryName, string subFolder = null)
        {
            UpdateSharepointFiles(_getSharepointPaths(siteUrl, libraryName, subFolder));
        }

        public static void UpdateSharepointFiles(IEnumerable<string> excelSharepointFilePaths)
        {
            Excel.Application excelApp = null;
            try
            {

                excelApp = new Excel.Application();
                _log.InfoFormat("Start of update process. Excel app initializated. {0} files to fetch:\n", 
                    excelSharepointFilePaths.Count(),
                    string.Join("\n",excelSharepointFilePaths));

                foreach(string path in excelSharepointFilePaths)
                {
                    try
                    {
                        _updateSharepointFile(excelApp, path);
                    }
                    catch (Exception ex)
                    {
                        string err = string.Format("Error occured for file {0}", path);
                        _log.Error(err, ex);
                    }

                }
            }
            finally
            {
                if (excelApp != null)
                {
                    excelApp.Quit();
                    _log.InfoFormat("End of update process. Excel app closed.");
                    excelApp = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static void _updateSharepointFile(Excel.Application excelApp, string excelSharepointFilePath)
        {
            string excelLocalWorkBookName = null;
            try
            {
                excelApp.Visible = false;
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(excelSharepointFilePath,
                    0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);

                excelWorkbook.RefreshAll();
                System.Threading.Thread.Sleep(5000);

                excelLocalWorkBookName = "temp_" + excelWorkbook.Name;

                string path = Path.Combine(Directory.GetCurrentDirectory(), excelLocalWorkBookName);
                excelWorkbook.SaveAs(path, ConflictResolution: Excel.XlSaveConflictResolution.xlLocalSessionChanges);
                
                excelWorkbook.Close(true);
                _publishWorkbook(path, excelSharepointFilePath);
                //excelApp.CalculateUntilAsyncQueriesDone();


            }
            finally
            {
                if (excelLocalWorkBookName != null)
                {
                    if (System.IO.File.Exists(excelLocalWorkBookName))
                    {
                        System.IO.File.Delete(excelLocalWorkBookName);
                    }
                    else
                    {
                        throw new Exception("Weird! We have temp file, but can not delete it, because it was not found");
                    }
                }

            }
        }

        private static void _publishWorkbook(string LocalPath, string SharePointPath)
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
                using (FileStream fsWorkbook = System.IO.File.Open(LocalPath,
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

        private static List<string> _getSharepointPaths (string siteUrl, string libraryName, string subFolder = null)
        {

            ClientContext context = new ClientContext(siteUrl);
            Web site = context.Web;
            context.Load(context.Web, w => w.ServerRelativeUrl);
            context.ExecuteQuery();

            List xlsList = site.Lists.GetByTitle(libraryName);
            CamlQuery caml = new CamlQuery();
            caml.ViewXml = "<View Scope=\"Recursive\"><Query><Where>";
            if (subFolder != null)
            {
                caml.ViewXml += "<And>";
            }
            caml.ViewXml += "<Eq><FieldRef Name=\"File_x0020_Type\"/><Value Type=\"Text\">xlsx</Value></Eq>";
            if (subFolder != null)
            {
                caml.ViewXml += "<Contains><FieldRef Name=\"FileDirRef\"/><Value Type=\"Text\">" + subFolder + "</Value></Contains>";
                caml.ViewXml += "</And>";
            }
            caml.ViewXml += "</Where></Query></View>";
            context.Load(xlsList);
            context.ExecuteQuery();
            var listItemCol = xlsList.GetItems(caml);
            context.Load(listItemCol);
            context.ExecuteQuery();

            List<string> result = new List<string>();

            // awfull, but for some reasons almost all linq methods throw System.NotSupportedException on this collection. Even ToList()
            foreach (ListItem item in listItemCol)
            {
                if (site.ServerRelativeUrl != "/")
                {
                    result.Add(string.Format("{0}/{1}",
                        siteUrl,
                        item["FileRef"].ToString().Replace(site.ServerRelativeUrl, "")));
                }
                else
                {
                    result.Add(string.Format("{0}/{1}",
                        siteUrl,
                        item["FileRef"].ToString()));
                }

            }
            return result;
            
        }
    }

}
