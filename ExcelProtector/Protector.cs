using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using Alphaleonis.Win32.Filesystem;
using Microsoft.Office.Interop.Excel;
using NLog;

namespace ExcelProtector
{
    public class Protector
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// Iterates over all files in the target directory applying protection 
        /// to the structure and worksheets of Excel workbooks with the provided
        /// file extensions.
        /// </summary>
        /// <param name="dir">The target directory to search.</param>
        /// <param name="extensions">The file extensions to search for.</param>
        /// <param name="password">The password applied to the Excel file.</param>
        /// <param name="token">The cancellation token used to check if the process should be cancelled.</param>
        public void ProtectFiles(DirectoryInfo dir, string[] extensions, string password, CancellationToken token)
        {
            logger.Debug("Preparing to protect files.");
            string searchPattern = Properties.Settings.Default.RegexFileSearchPattern
                .Replace("{extensions}", String.Join("|", extensions.OrderByDescending(e => e)));
            logger.Debug($"Search Pattern: {searchPattern}");
            Regex regex = new Regex(searchPattern, RegexOptions.IgnoreCase);
            logger.Debug($"RegEx Options: {regex.Options}");
            logger.Debug($"Target Folder: {dir?.FullName}");
            IEnumerable<FileInfo> files = dir.GetFiles("*", System.IO.SearchOption.AllDirectories)
                .Where(file => regex.IsMatch(file.Name));
            int fileCount = 0;
            logger.Debug($"{files.Count()} Excel files have been identified.");

            foreach (FileInfo file in files)
            {
                if (token.IsCancellationRequested)
                {
                    OnError(new Exception("The process has been cancelled."));
                    break;
                }
                else
                {
                    try
                    {
                        Protect(file, password);
                        OnFileProtected(file);
                    }
                    catch (Exception ex)
                    {
                        OnError(ex);
                    }
                }

                fileCount++;
                int progress = (int)((double)fileCount / (double)files.Count() * 100f);
                OnReportProgress(progress);
            }
        }

        /// <summary>
        /// Applies worksheet and structure protection to a target Excel file.
        /// </summary>
        /// <param name="excelFile">The target excel file.</param>
        /// <param name="password">The password used to apply protection.</param>
        public static void Protect(FileInfo excelFile, string password)
        {
            Application app = new Application();
            app.Visible = false;
                        
            try
            {
                var workbook = app.Workbooks.Open(excelFile.FullName);

                if (workbook.ProtectStructure)
                {
                    throw new Exception($"The workbook's ({excelFile.FullName}) structure is already protected.");
                }
                else if (workbook.ProtectWindows)
                {
                    throw new Exception($"The workbook's ({excelFile.FullName}) windows are protected.");
                }
                else if (workbook.HasPassword)
                {
                    throw new Exception($"The workbook ({excelFile.FullName}) has a password.");
                }

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (WorksheetIsProtected(sheet))
                    {
                        throw new Exception($"The worksheet {sheet.Name} in {excelFile.FullName} is protected.");
                    }

                    sheet.Protect(password);
                }

                workbook.Protect(password, true);
                workbook.Save();
            }
            finally
            {
                app.Quit();
            }            
        }

        /// <summary>
        /// Evaluates if a target worksheet is protected.
        /// </summary>
        /// <param name="sheet">The target sheet to evaluate.</param>
        /// <returns>Returns true if the sheet is protected, otherwise false.</returns>
        private static bool WorksheetIsProtected(Worksheet sheet)
        {
            bool result;

            try
            {
                sheet.Unprotect(String.Empty);
                result = false;
            }
            catch
            {
                result = true;
            }

            return result;
        }

        /// <summary>
        /// This event occurs when the 
        /// <see cref="ProtectFiles(DirectoryInfo, string[], string, CancellationToken)"/> 
        /// method reports it's progress.
        /// </summary>
        public event EventHandler<int> ReportProgress;

        /// <summary>
        /// This event occurs when the 
        /// <see cref="ProtectFiles(DirectoryInfo, string[], string, CancellationToken)"/>
        /// method reports that a file has been protected.
        /// </summary>
        public event EventHandler<FileInfo> FileProtected;

        /// <summary>
        /// This event occurs when the 
        /// <see cref="ProtectFiles(DirectoryInfo, string[], string, CancellationToken)"/> 
        /// method reports that an exception has occured.
        /// </summary>
        public event EventHandler<Exception> Error;
        
        protected virtual void OnReportProgress(int progress)
        {
            ReportProgress(this, progress);
        }

        protected virtual void OnFileProtected(FileInfo file)
        {
            FileProtected?.Invoke(this, file);
        }

        protected virtual void OnError(Exception ex)
        {
            Error?.Invoke(this, ex);
        }
    }
}
