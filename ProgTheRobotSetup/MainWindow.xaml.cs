using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;
using IWshRuntimeLibrary;
using System;
using System.IO;
using System.Net;
using System.Windows;
using Octokit;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Reflection;
using System.Threading;
using System.Diagnostics;

namespace ProgTheRobotSetup
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        #region ButtonEvents

        private void BackButton(object sender, RoutedEventArgs e)
        {
            ShowPreviousGrid();
        }

        private void StartInstall(object sender, RoutedEventArgs e)
        {
            ShowNextGrid();
            ShowReleaseTag();
        }

        /// <summary>
        /// Remove the program
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Uninstall(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Do you really want to uninstall this program ?", "Uninstall", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {

                if (IsFileLocked(new FileInfo(System.IO.Path.Combine(INSTALL_PATH, "Prog the robot.exe"))))
                {
                    MessageBox.Show("Please close Prog The Robot First", "Uninstall", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                RemoveStartMenuEntry();

                var procId = Process.GetCurrentProcess().Id;

                //remove the regestry keys

                StreamWriter stream = System.IO.File.CreateText(UNINSTALLBAT_PATH);
                stream.WriteLine($"Taskkill /F /PID {procId}");
                stream.WriteLine($"rmdir /S /Q \"{INSTALL_PATH}\"");
                stream.WriteLine("powershell -Command \"& {Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.MessageBox]::Show('Prog The Robot was successfully removed', 'Uninstaller', 'OK', [System.Windows.Forms.MessageBoxIcon]::Information);}\"");
                stream.WriteLine($"del /Q \"{UNINSTALLBAT_PATH}\"");
                stream.Close();
                Process process = new Process();
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.FileName = UNINSTALLBAT_PATH;
                process.Start();
                System.Windows.Application.Current.Shutdown();
            }

        }
        private void InstallSettingsNext(object sender, RoutedEventArgs e)
        {
            ShowNextGrid();
            Install();
        }
        #endregion

        // install constants
        const string INSTALL_PATH = @"C:\Program Files\ProgTheRobot";
        const string UNINSTALLBAT_PATH = @"C:\Program Files\uninstall46746234645.bat";
        const string TEMP_PATH = @"C:\Program Files\ProgTheRobot\temp\";
        const string DL_FILE_NAME = "progtherobot";
        const string PROGRAMS_PATH = @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\ProgTheRobot";

        enum GridPanel
        {
            Main = 0,
            InstallConf,
            Download,
            Finished
        }


        private Dictionary<GridPanel, Grid> gridShowOrder = new Dictionary<GridPanel, Grid>();
        private GridPanel currentPanel = GridPanel.Main;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            DownloadGrid.Visibility = Visibility.Hidden;
            InstallSettingsGrid.Visibility = Visibility.Hidden;
            FinishedGrid.Visibility = Visibility.Hidden;

            gridShowOrder.Add(GridPanel.Main, MainGrid);
            gridShowOrder.Add(GridPanel.InstallConf, InstallSettingsGrid);
            gridShowOrder.Add(GridPanel.Download, DownloadGrid);
            gridShowOrder.Add(GridPanel.Finished, FinishedGrid);

            VerLabel.Visibility = Visibility.Hidden;
        }

        private void Install()
        {
            CreateDir(new string[] { INSTALL_PATH, TEMP_PATH, PROGRAMS_PATH });
            string path = Process.GetCurrentProcess().MainModule.FileName;
            string fileName = System.IO.Path.GetFileName(path);
            if (!System.IO.File.Exists(System.IO.Path.Combine(INSTALL_PATH, fileName)))
            {
                System.IO.File.Copy(path, System.IO.Path.Combine(INSTALL_PATH, fileName), true);
            }
            GetDownloadLinks((url) =>
            {
                DownloadFromUrl(
                new Uri(url),
                System.IO.Path.Combine(TEMP_PATH, DL_FILE_NAME),
                () =>
                {
                    FileStream stream = System.IO.File.OpenRead(System.IO.Path.Combine(TEMP_PATH, DL_FILE_NAME));
                    UnzipFromStream(stream, INSTALL_PATH);
                    CreateStartMenuEntry(System.IO.Path.Combine(INSTALL_PATH, "Prog the robot.exe"), "Prog The Robot.lnk", "Launch Prog The Robot");
                    RemoveDir(new string[] { TEMP_PATH });
                    CreateStartMenuEntry(System.IO.Path.Combine(INSTALL_PATH, fileName), "Prog The Robot installer.lnk", "Install, update and uninstall Prog The Robot");

                    FileAssociation fileAssociation = new FileAssociation("Prog The Robot");
                    fileAssociation.SetExtension(".pr", System.IO.Path.Combine(INSTALL_PATH, "Prog the robot.exe"));

                    ShowNextGrid();
                });
            });
        }

        private void FileAssociation()
        {

        }

        private void ShowMain()
        {
            gridShowOrder[currentPanel].Visibility = Visibility.Hidden;
            int test = (int)GridPanel.Main;
            currentPanel = Enum.Parse<GridPanel>((test).ToString());
            gridShowOrder[currentPanel].Visibility = Visibility.Visible;
        }

        /// <summary>
        /// Show the next grid
        /// </summary>
        private void ShowNextGrid()
        {
            try
            {
                gridShowOrder[currentPanel].Visibility = Visibility.Hidden;
                int test = (int)currentPanel;
                currentPanel = Enum.Parse<GridPanel>((test + 1).ToString());
                gridShowOrder[currentPanel].Visibility = Visibility.Visible;
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// Show the previous grid
        /// </summary>
        private void ShowPreviousGrid()
        {
            try
            {
                gridShowOrder[currentPanel].Visibility = Visibility.Hidden;
                int test = (int)currentPanel;
                currentPanel = Enum.Parse<GridPanel>((test - 1).ToString());
                gridShowOrder[currentPanel].Visibility = Visibility.Visible;
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// <see cref="https://www.codercream.com/check-file-use-c-code/"/>
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        protected virtual bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(System.IO.FileMode.Open, FileAccess.Write, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }


        private async void DownloadFromUrl(Uri ressource, string savePath, Action endCallback)
        {
            WebClient webClient = new WebClient();
            webClient.DownloadProgressChanged += onProgressChanged;
            await webClient.DownloadFileTaskAsync(ressource, savePath);
            endCallback?.Invoke();
        }
        private void onProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            DownLoadProgress.Value = e.ProgressPercentage;
        }

        private void CreateDir(string[] dirs)
        {
            foreach (string path in dirs)
            {
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
            }
        }

        private void RemoveDir(string[] dirs)
        {
            foreach (string path in dirs)
            {
                if (Directory.Exists(path))
                    Directory.Delete(path, true);
            }
        }


        /// <summary>
        /// Create a start menu entry<see cref="https://morgantechspace.com/2015/01/create-start-menu-shortcut-all-programs-csharp.html"/>
        /// </summary>
        /// <param name="targetPath">The path to the executable</param>
        private void CreateStartMenuEntry(string targetPath, string linkName, string description)
        {
            string appLink = System.IO.Path.Combine(PROGRAMS_PATH, linkName);
            IWshShortcut shortcut = (IWshShortcut)new WshShell().CreateShortcut(appLink);
            shortcut.TargetPath = targetPath;
            shortcut.Description = description;
            shortcut.Save();
        }

        /// <summary>
        /// Remove the start menu entry
        /// </summary>
        private void RemoveStartMenuEntry()
        {
            Directory.Delete(PROGRAMS_PATH, true);
        }

        /// <summary>
        /// code from <see cref="https://github.com/icsharpcode/SharpZipLib/wiki/Unpack-a-zip-using-ZipInputStream"/>
        /// </summary>
        /// <param name="zipStream">A stream with a zip</param>
        /// <param name="outFolder">The folder where the unzip files will be saved</param>
        private void UnzipFromStream(Stream zipStream, string outFolder)
        {
            var zipInputStream = new ZipInputStream(zipStream);
            while (zipInputStream.GetNextEntry() is ZipEntry zipEntry)
            {
                var entryFileName = zipEntry.Name;
                // To remove the folder from the entry:
                //var entryFileName = Path.GetFileName(entryFileName);
                // Optionally match entrynames against a selection list here
                // to skip as desired.
                // The unpacked length is available in the zipEntry.Size property.

                // 4K is optimum
                var buffer = new byte[4096];

                // Manipulate the output filename here as desired.
                var fullZipToPath = System.IO.Path.Combine(outFolder, entryFileName);
                var directoryName = System.IO.Path.GetDirectoryName(fullZipToPath);
                if (directoryName.Length > 0)
                    Directory.CreateDirectory(directoryName);

                // Skip directory entry
                if (System.IO.Path.GetFileName(fullZipToPath).Length == 0)
                {
                    continue;
                }

                // Unzip file in buffered chunks. This is just as fast as unpacking
                // to a buffer the full size of the file, but does not waste memory.
                // The "using" will close the stream even if an exception occurs.
                using (FileStream streamWriter = System.IO.File.Create(fullZipToPath))
                {
                    StreamUtils.Copy(zipInputStream, streamWriter, buffer);
                }
            }
        }


        /// <summary>
        /// <see cref="https://stackoverflow.com/questions/25678690/how-can-i-check-github-releases-in-c"/>
        /// </summary>
        private async void GetDownloadLinks(Action<string> endCallBack)
        {
            try
            {
                GitHubClient client = new GitHubClient(new ProductHeaderValue("ProgTheRobotSetup"));
                IReadOnlyList<Release> releases = await client.Repository.Release.GetAll("jolanaklin", "progtherobot");
                var latest = releases[0];
                string url = Array.Find<ReleaseAsset>(latest.Assets.ToArray(), x => x.Name == "ProgTheRobot.zip").BrowserDownloadUrl;
                endCallBack?.Invoke(url);
            }
            catch (Exception) { MessageBox.Show("Can't connect to GitHub", "Connecting to GitHub", MessageBoxButton.OK, MessageBoxImage.Error); ShowMain(); }
        }

        private async void ShowReleaseTag()
        {
            try
            {
                GitHubClient client = new GitHubClient(new ProductHeaderValue("ProgTheRobotSetup"));
                IReadOnlyList<Release> releases = await client.Repository.Release.GetAll("jolanaklin", "progtherobot");
                var latest = releases[0];
                VerLabel.Content = latest.TagName;
                SolidColorBrush solidColor = null;
                if (latest.Prerelease)
                {
                    solidColor = new SolidColorBrush(Color.FromRgb(Convert.ToByte(255), Convert.ToByte(133), Convert.ToByte(0)));
                }
                else
                {
                    solidColor = new SolidColorBrush(Color.FromRgb(Convert.ToByte(0), Convert.ToByte(128), Convert.ToByte(0)));
                }
                VerLabel.Background = solidColor;
                VerLabel.Visibility = Visibility.Visible;
            }
            catch (Exception) { MessageBox.Show("Can't connect to GitHub", "Connecting to GitHub", MessageBoxButton.OK, MessageBoxImage.Error); ShowMain(); }
        }
    }
}
