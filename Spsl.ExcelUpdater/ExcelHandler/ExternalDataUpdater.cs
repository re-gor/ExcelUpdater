using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using Microsoft.SharePoint.Client;
using log4net;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelHandler
{
    /// <summary>
    /// This class updates external data connections in excel and update data in excel file
    /// </summary>
    public static class ExternalDataUpdater
    {
        private const int _waitUpdateTimeout = 5000;
        private const int _waitDeleteErrorTimeout = 10000;
        private static List<Task> _deleteTasks = new List<Task>();

        private static readonly ILog _log = LogManager.GetLogger(typeof(ExternalDataUpdater));

        /// <summary>
        /// <para>Recursively get urls for all xlsx files in choosen directory and subdirectory</para>
        /// <para>For each file: Open it from Sharepoint site, refresh connections, </para>
        /// <para>save it locally, publish back it to Sharepoint and delete its local copy</para>
        /// </summary>
        /// <param name="siteUrl">Url of site, where library is located</param>
        /// <param name="libraryName">Name of library ("Shared documents", for example)</param>
        /// <param name="subFolder">Specific subfolder in library. Function will add "Contains" operator in caml query</param>
        public static void UpdateSharepointFiles(string siteUrl, string libraryName, string subFolder = null)
        {
            UpdateSharepointFiles(_getSharepointPaths(siteUrl, libraryName, subFolder));
            Task.WaitAll(_deleteTasks.ToArray());
        }

        /// <summary>
        /// Updating of set of files. 
        /// </summary>
        /// <param name="excelSharepointFilePaths"></param>
        public static void UpdateSharepointFiles(IEnumerable<string> excelSharepointFilePaths)
        {
            Excel.Application excelApp = null;
            try
            {

                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;

                _log.InfoFormat("Start of update process. Excel app initializated. {0} files to fetch:\r\n{1}", 
                    excelSharepointFilePaths.Count(),
                    string.Join("\r\n", excelSharepointFilePaths));

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

        /// <summary>
        /// Open excel file from Sharepoint site, refresh connections, save it locally, publish back it to Sharepoint and delete its local copy
        /// </summary>
        /// <param name="excelApp">Instance of excel application for file opening</param>
        /// <param name="excelSharepointFilePath">Url of file</param>
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
                System.Threading.Thread.Sleep(_waitUpdateTimeout);

                excelLocalWorkBookName = "temp_" + excelWorkbook.Name;

                string path = Path.Combine(Directory.GetCurrentDirectory(), excelLocalWorkBookName);
                excelWorkbook.SaveAs(path, ConflictResolution: Excel.XlSaveConflictResolution.xlLocalSessionChanges);
                
                excelWorkbook.Close(true);
                _publishWorkbook(path, excelSharepointFilePath);
            }
            finally
            {
                if (excelLocalWorkBookName != null)
                {
                    if (System.IO.File.Exists(excelLocalWorkBookName))
                    {
                        try {
                            System.IO.File.Delete(excelLocalWorkBookName);
                        }
                        catch (System.IO.IOException ex)
                        {
                            string message = string.Format("Can not delete file {0}. Will try to delete it async after {1} sec. Error: ", 
                                excelLocalWorkBookName, _waitDeleteErrorTimeout);
                            _log.Error(message, ex);
                            _deleteTasks.Add(_deleteErrorHandler(excelLocalWorkBookName));
                        }
                    }
                    else
                    {
                        throw new Exception("Weird! We have temp file, but can not delete it, because it was not found");
                    }
                }

            }
        }

        /// <summary>
        /// If delete opeartion of local file fails, this one will try to delete file in parallel task
        /// </summary>
        /// <param name="excelLocalWorkBookName">Name of file to delete</param>
        private static Task _deleteErrorHandler(string excelLocalWorkBookName)
        {
            var t = new Task(() => {
                for (int i = 1; i <= 5; ++i)
                {
                    try
                    {
                        System.Threading.Thread.Sleep(_waitDeleteErrorTimeout);
                        System.IO.File.Delete(excelLocalWorkBookName);

                        break;
                    }
                    catch (System.IO.IOException ex)
                    {
                        string message = string.Format("Can not delete file {0}. ASync attempt {1} of {2} failed. Will try again after {2} sec. Error: ", 
                            excelLocalWorkBookName, i, 5, _waitDeleteErrorTimeout);
                        _log.Fatal(message, ex);
                    }
                }
            });

            t.Start();
            return t;
        }

        /// <summary>
        /// Publish xlsx file to Sharepoint Server
        /// </summary>
        /// <param name="LocalPath"></param>
        /// <param name="SharePointPath"></param>
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

        /// <summary>
        /// Recursively get urls for all xlsx files in choosen directory and subdirectory
        /// </summary>
        /// <param name="siteUrl">Url of site, where library is located</param>
        /// <param name="libraryName">Name of library ("Shared documents", for example)</param>
        /// <param name="subFolder">Specific subfolder in library. Function will add "Contains" operator in caml query</param>
        /// <returns></returns>
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
