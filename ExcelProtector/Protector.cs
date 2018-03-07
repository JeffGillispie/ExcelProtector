using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using Alphaleonis.Win32.Filesystem;
using Microsoft.Office.Interop.Excel;

namespace ExcelProtector
{
    public class Protector
    {
        public void ProtectFiles(DirectoryInfo dir, string[] extensions, string password, CancellationToken token)
        {            
            string searchPattern = Properties.Settings.Default.RegexFileSearchPattern
                .Replace("{extensions}", String.Join("|", extensions.OrderByDescending(e => e)));
            Regex regex = new Regex(searchPattern, RegexOptions.IgnoreCase);
            IEnumerable<FileInfo> files = dir.GetFiles("*", System.IO.SearchOption.AllDirectories)
                .Where(file => regex.IsMatch(file.Name));
            int fileCount = 0;

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
                sheet.Protect(password);
            }

            workbook.Save();
            app.Quit();
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
