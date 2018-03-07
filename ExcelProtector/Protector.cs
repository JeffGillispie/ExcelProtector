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

        public static void Protect(FileInfo excelFile, string password)
        {
            Application app = new Application();
            app.Visible = false;

            var workbook = app.Workbooks.Open(excelFile.FullName);

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (WorksheetIsProtected(sheet))
                {
                    throw new Exception($"The worksheet {sheet.Name} in {excelFile.FullName} is protected.");
                }

                sheet.Protect(password);
            }

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

            workbook.Protect(password, true);
            workbook.Save();
            app.Quit();
        }

        private static bool WorksheetIsProtected(Worksheet sheet)
        {
            bool result;

            try
            {
                sheet.Unprotect(String.Empty);
                result = true;
            }
            catch
            {
                result = false;
            }

            return result;
        }

        public event EventHandler<int> ReportProgress;
        public event EventHandler<FileInfo> FileProtected;
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
