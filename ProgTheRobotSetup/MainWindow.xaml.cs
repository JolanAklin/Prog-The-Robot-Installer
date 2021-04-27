using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;
using IWshRuntimeLibrary;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ProgTheRobotSetup
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            DownloadGrid.Visibility = Visibility.Hidden;
        }

        private void StartInstall(object sender, RoutedEventArgs e)
        {
            MainGrid.Visibility = Visibility.Hidden;
            DownloadGrid.Visibility = Visibility.Visible;

            string installPath = $@"C:\Program Files\ProgTheRobot";
            string path = $@"C:\Program Files\ProgTheRobot\temp\";
            string dlFileName = "progtherobot";
            if (Directory.Exists(installPath))
                Directory.Delete(installPath);
            Directory.CreateDirectory(installPath);
            if (Directory.Exists(path))
                Directory.Delete(path);
            Directory.CreateDirectory(path);
            DownloadFromUrl(
                new Uri("https://github.com/JolanAklin/ProgTheRobot/releases/download/v1.0a/Prog_the_robot_v1-0a.zip"),
                System.IO.Path.Combine(path, dlFileName), 
                () => {
                    FileStream stream = System.IO.File.OpenRead(System.IO.Path.Combine(path, dlFileName));
                    UnzipFromStream(stream, installPath);
                    CreateStartMenuEntry(System.IO.Path.Combine(installPath, "prog_the_robot/Prog the robot.exe"));
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


        /// <summary>
        /// Create a start menu entry<see cref="https://morgantechspace.com/2015/01/create-start-menu-shortcut-all-programs-csharp.html"/>
        /// </summary>
        /// <param name="targetPath">The path to the executable</param>
        private void CreateStartMenuEntry(string targetPath)
        {
            string programs_path = @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs";
            string settingsLink = System.IO.Path.Combine(programs_path, "ProgTheRobot.lnk");
            IWshShortcut shortcut = (IWshShortcut)new WshShell().CreateShortcut(settingsLink);
            shortcut.TargetPath = targetPath;
            //shortcut.IconLocation = @"C:\Program FilesMorganTechSpacesettings.ico";
            //shortcut.Arguments = "arg1 arg2";
            shortcut.Description = "Launch Prog The Robot";
            shortcut.Save();
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
    }
}
