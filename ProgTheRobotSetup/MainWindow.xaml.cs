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

namespace ProgTheRobotSetup
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string installPath = $@"C:\Program Files\ProgTheRobot";
        string path = $@"C:\Program Files\ProgTheRobot\temp\";
        string dlFileName = "progtherobot";
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            DownloadGrid.Visibility = Visibility.Hidden;
        }
        private void Uninstall(object sender, RoutedEventArgs e)
        {
            if(MessageBox.Show("Do you really want to uninstall this program ?", "Uninstall", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                RemoveInstallDir(new string[] { installPath, path });
                RemoveStartMenuEntry();
                // need to remove files in locallow and the file association
                MessageBox.Show("The programm has been uninstalled properly");
            }
        }

        private void StartInstall(object sender, RoutedEventArgs e)
        {
            MainGrid.Visibility = Visibility.Hidden;
            DownloadGrid.Visibility = Visibility.Visible;



            RemoveInstallDir(new string[] { installPath, path });
            CreateInstallDir(new string[] { installPath, path });

            GetDownloadLinks((url) =>
            { 
                DownloadFromUrl(
                new Uri(url),
                System.IO.Path.Combine(path, dlFileName),
                () =>
                {
                    FileStream stream = System.IO.File.OpenRead(System.IO.Path.Combine(path, dlFileName));
                    UnzipFromStream(stream, installPath);
                    CreateStartMenuEntry(System.IO.Path.Combine(installPath, "Prog the robot.exe"));
                    MessageBox.Show("Finished");
                });
            });
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

        private void CreateInstallDir(string[] dirs)
        {
            foreach (string path in dirs)
            {
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
            }
        }

        private void RemoveInstallDir(string[] dirs)
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
        private void CreateStartMenuEntry(string targetPath)
        {
            string programs_path = @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs";
            string settingsLink = System.IO.Path.Combine(programs_path, "Prog The Robot.lnk");
            IWshShortcut shortcut = (IWshShortcut)new WshShell().CreateShortcut(settingsLink);
            shortcut.TargetPath = targetPath;
            shortcut.Description = "Launch Prog The Robot";
            shortcut.Save();
        }

        /// <summary>
        /// Remove the start menu entry
        /// </summary>
        private void RemoveStartMenuEntry()
        {
            string programs_path = @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs";
            System.IO.File.Delete(System.IO.Path.Combine(programs_path, "Prog The Robot.lnk"));
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
            GitHubClient client = new GitHubClient(new ProductHeaderValue("ProgTheRobotSetup"));
            IReadOnlyList<Release> releases = await client.Repository.Release.GetAll("jolanaklin", "progtherobot");
            var latest = releases[0];
            string url = Array.Find<ReleaseAsset>(latest.Assets.ToArray(), x => x.Name == "ProgTheRobot.zip").BrowserDownloadUrl;
            endCallBack?.Invoke(url);
        }

    }
}
