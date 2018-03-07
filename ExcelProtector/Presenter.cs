using System;
using System.ComponentModel;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Alphaleonis.Win32.Filesystem;
using NLog;

namespace ExcelProtector
{
    /// <summary>
    /// The view model for the main window.
    /// </summary>
    public class Presenter : ObservableObject
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        private DirectoryInfo targetFolder;        
        private bool isEnabled = true;
        private bool isWorking = true;
        private int workerProgress = 0;
        private int errorCount = 0;
        private int protectedCount = 0;        
        private BackgroundWorker worker = new BackgroundWorker() {
            WorkerReportsProgress = true,
            WorkerSupportsCancellation = true
        };
        
        /// <summary>
        /// The path to the folder containing the target Excel files.
        /// </summary>
        public string TargetFolderPath
        {
            get
            {
                return targetFolder?.FullName ?? String.Empty;
            }

            set
            {
                logger.Trace("Setting target folder path.");

                if (targetFolder == null || !targetFolder.FullName.Equals(value))
                {
                    logger.Trace($"Target Folder Path: {value}");
                    targetFolder = new DirectoryInfo(value);
                    OnPropertyChanged(nameof(TargetFolderPath));
                }
            }
        }

        /// <summary>
        /// Informational text about the process to display to the user.
        /// </summary>
        public string InfoText
        {
            get
            {
                StringBuilder builder = new StringBuilder();
                var extensions = Properties.Settings.Default.TargetedFileExtensions.Split('|');
                var extList = String.Join(", ", extensions);
                builder
                    .Append("The target folder and sub-folders will be searched for files ")
                    .Append($"with the following extensions ({extList}). ")                    
                    .Append("The resulting Excel files will have all constituent worksheets ")
                    .Append("protected and saved with the provided password.");
                return builder.ToString();
            }
        }
        
        /// <summary>
        /// Returns false when the UI text boxes and ok button are disabled otherwise true.
        /// </summary>
        public bool IsEnabled
        {
            get
            {
                return isEnabled;
            }

            set
            {
                if (isEnabled != value)
                {
                    isEnabled = value;
                    OnPropertyChanged(nameof(IsEnabled));
                }
            }
        }

        /// <summary>
        /// Returns true when the worker is working, otherwise false.
        /// </summary>
        public bool IsWorking
        {
            get
            {
                return isWorking;
            }

            set
            {
                if (isWorking != value)
                {
                    isWorking = value;
                    OnPropertyChanged(nameof(IsWorking));
                }
            }
        }

        /// <summary>
        /// Returns the progress percentage of the background worker.
        /// </summary>
        public int WorkerProgress
        {
            get
            {
                return workerProgress;
            }

            set
            {
                if (!workerProgress.Equals(value))
                {
                    workerProgress = value;
                    OnPropertyChanged(nameof(WorkerProgress));
                }
            }
        }
        
        /// <summary>
        /// The execute command for the background worker.
        /// </summary>
        public ICommand ExecuteCommand
        {
            get
            {
                var command = new DelegateCommand((param) => Execute(param));
                return command;
            }
        }

        /// <summary>
        ///  The cancel command for the background worker.
        /// </summary>
        public ICommand CancelCommand
        {
            get
            {
                var command = new DelegateCommand((param) => Cancel());
                return command;
            }
        }

        /// <summary>
        /// The get target folder command.
        /// </summary>
        public ICommand GetTargetFolderCommand
        {
            get
            {
                var command = new DelegateCommand((param) => GetTargetFolder());
                return command;
            }
        }

        /// <summary>
        /// Execute the background worker and the protection process.
        /// </summary>
        /// <param name="parameter"></param>
        private void Execute(object parameter)
        {
            logger.Debug("Preparing background worker.");
            IsEnabled = false;
            WorkerProgress = 0;
            errorCount = 0;
            protectedCount = 0;
            worker.DoWork += Work;
            worker.ProgressChanged += WorkProgress;
            worker.RunWorkerCompleted += WorkComplete;
            var extensions = Properties.Settings.Default.TargetedFileExtensions.Split('|');
            var passwordBox = parameter as PasswordBox;
            var workArgs = new WorkArgs() {
                Folder = targetFolder,
                Extensions = extensions,
                PasswordBox = passwordBox
            };

            logger.Debug("Initiating worker.");
            worker.RunWorkerAsync(workArgs);
        }

        /// <summary>
        /// Requests cancellation of the background worker.
        /// </summary>
        private void Cancel()
        {
            worker.CancelAsync();
        }

        /// <summary>
        /// Launches a folder browser dialog and prompts the user for a target folder.
        /// </summary>
        private void GetTargetFolder()
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TargetFolderPath = dialog.SelectedPath;
                }
            }
        }

        /// <summary>
        /// Handles the do work event.
        /// </summary>
        /// <param name="sender">Work event sender.</param>
        /// <param name="e">Work event arguments.</param>
        private void Work(object sender, DoWorkEventArgs e)
        {
            logger.Debug("Preparing protector.");
            var args = e.Argument as WorkArgs;                        
            CancellationTokenSource cts = new CancellationTokenSource();
            Protector protector = new Protector();
            protector.FileProtected += (fpSender, fi) => worker.ReportProgress(0, fi);
            protector.Error += (exSender, ex) => worker.ReportProgress(0, ex);

            protector.ReportProgress += (rpSender, progress) => {
                worker.ReportProgress(progress);

                if (worker.CancellationPending)
                    cts.Cancel();
            };

            logger.Debug("Starting the protection process.");

            try
            {
                protector.ProtectFiles(
                    args.Folder,
                    args.Extensions,
                    args.PasswordBox.Password,
                    cts.Token);
            }
            catch (Exception ex)
            {
                logger.Error(ex);
                e.Result = ex;
            }
        }

        /// <summary>
        /// Handles the background worker progress changed events.
        /// </summary>
        /// <param name="sender">The work progress sender.</param>
        /// <param name="e">Progress changed event arguments.</param>
        private void WorkProgress(object sender, ProgressChangedEventArgs e)
        {
            if (e.UserState != null && e.UserState.GetType().Equals(typeof(FileInfo)))
            {
                var file = e.UserState as FileInfo;
                logger.Info($"File Protected: {file.FullName}");
                protectedCount++;
            }
            else if (e.UserState != null && e.UserState.GetType().Equals(typeof(Exception)))
            {
                var ex = e.UserState as Exception;
                logger.Error(ex);
                errorCount++;
            }
            else
            {
                WorkerProgress = e.ProgressPercentage;
            }
        }

        /// <summary>
        /// Handles the background worker completed event.
        /// </summary>
        /// <param name="sender">The worker completed sender.</param>
        /// <param name="e">The worker completed event arguments.</param>
        private void WorkComplete(object sender, RunWorkerCompletedEventArgs e)
        {            
            IsEnabled = true;
            IsWorking = false;

            if (e.Result != null && e.Result.GetType().Equals(typeof(Exception)))
            {
                var ex = e.Result as Exception;
                MessageBox.Show($"There was a problem during the protection process. See the log file for details. {ex.Message}");
            }
            else
            {
                WorkerProgress = 100;
                logger.Debug($"The process has completed with {errorCount} errors.");
                MessageBox.Show($"The process has completed with {protectedCount} files protected and {errorCount} errors. See the log file for details.");
            }
        }

        private class WorkArgs
        {
            public DirectoryInfo Folder { get; set; }
            public string[] Extensions { get; set; }
            public PasswordBox PasswordBox { get; set; }
        }
    }
}
