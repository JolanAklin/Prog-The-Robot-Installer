using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;
using IWshRuntimeLibrary;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProgTheRobotSetup
{
    public class Install
    {
        /// <summary>
        /// Create dirs if they don't exist
        /// </summary>
        /// <param name="dirs">an array with the path of dirs to create</param>
        public static void CreateDir(string[] dirs)
        {
            foreach (string path in dirs)
            {
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
            }
        }

        /// <summary>
        /// remove dirs if they exist
        /// </summary>
        /// <param name="dirs">an array with the path of dirs to remove</param>
        public static void RemoveDir(string[] dirs)
        {
            foreach (string path in dirs)
            {
                if (Directory.Exists(path))
                    Directory.Delete(path, true);
            }
        }

        private string PROGRAMS_PATH = @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\ProgTheRobot";
        private string INSTALL_PATH = @"C:\Program Files\ProgTheRobot";
        private string UNINSTALLBAT_PATH = @"C:\Program Files\uninstall46746234645.bat";
        private string TEMP_PATH = @"C:\Program Files\ProgTheRobot\temp\";
        private string DL_FILE_NAME = "progtherobot";
        private string fileName = null;

        /// <summary>
        /// Install constructor
        /// </summary>
        /// <param name="programs_shortcut_path">Where the link to the program will be stored</param>
        /// <param name="install_path">Where the program is stored</param>
        /// <param name="temp_path">Where the temp download directory</param>
        /// <param name="dl_file_name">The name of the downloaded file</param>
        /// <param name="uninstall_path">Where the batch script for uninstall will be stored</param>
        public Install(string programs_shortcut_path, string install_path, string temp_path, string dl_file_name, string uninstall_path)
        {
            PROGRAMS_PATH = programs_shortcut_path;
            INSTALL_PATH = install_path;
            UNINSTALLBAT_PATH = uninstall_path;
            TEMP_PATH = temp_path;
            DL_FILE_NAME = dl_file_name;
        }

        /// <summary>
        /// Create the needed directory for the installation
        /// </summary>
        public void PreInstall()
        {
            CreateDir(new string[] { INSTALL_PATH, TEMP_PATH, PROGRAMS_PATH });

            // copy the installer if there is not already one in the install directory
            string path = Process.GetCurrentProcess().MainModule.FileName;
            fileName = System.IO.Path.GetFileName(path);
            if (!System.IO.File.Exists(System.IO.Path.Combine(INSTALL_PATH, fileName)))
            {
                System.IO.File.Copy(path, System.IO.Path.Combine(INSTALL_PATH, fileName), true);
            }
        }

        /// <summary>
        /// install the program, add start menu entry and add the file association
        /// </summary>
        /// <param name="CallBack">Called when the installation is finished</param>
        public async void InstallApp(GitHubReleaseFetcher.DownloadableFiles[] downloadedFiles, Action CallBack)
        {
            if(fileName == null)
            {
                PreInstall();
            }

            bool needToCreateProgTheRobotDevLink = false;
            bool needToCreateProgTheRobotLink = false;

            bool devInstall = false;
            bool soundInstall = false;

            foreach (GitHubReleaseFetcher.DownloadableFiles downloadedFile in downloadedFiles)
            {
                //unzip the downloaded Prog the robot release
                FileStream stream = System.IO.File.OpenRead(System.IO.Path.Combine(TEMP_PATH, downloadedFile.ToString()));
                string path = "";
                switch (downloadedFile)
                {
                    case GitHubReleaseFetcher.DownloadableFiles.ProgTheRobot:
                        path = INSTALL_PATH;
                        needToCreateProgTheRobotLink = true;
                        break;
                    case GitHubReleaseFetcher.DownloadableFiles.SoundPack:
                        path = Path.Combine(INSTALL_PATH, "Prog the robot_Data", "sounds");
                        soundInstall = true;
                        break;
                    case GitHubReleaseFetcher.DownloadableFiles.DemoPack:
                        path = $@"C:/Users/{Environment.UserName}/AppData/LocalLow/Jolan Aklin/Prog the robot/saves/";
                        break;
                    case GitHubReleaseFetcher.DownloadableFiles.ProgTheRobotDev:
                        path = Path.Combine(INSTALL_PATH, "dev");
                        devInstall = true;
                        needToCreateProgTheRobotDevLink = true;
                        break;
                }
                CreateDir(new string[] { path });
                Task t = new Task(() => { UnzipFromStream(stream, path); });
                t.Start();
                await t;
            }

            if(devInstall && soundInstall)
            {
                string devSound = Path.Combine(INSTALL_PATH, "dev", "Prog the robot_Data", "sounds");
                CreateDir(new string[] { devSound});
                DirectoryInfo directoryInfo = new DirectoryInfo(Path.Combine(INSTALL_PATH, "Prog the robot_Data", "sounds"));
                foreach (FileInfo file in directoryInfo.GetFiles())
                {
                    file.CopyTo(Path.Combine(devSound, file.Name));
                }
            }    

            RemoveDir(new string[] { TEMP_PATH });

            // create start menu entry for the program and the installer
            if(needToCreateProgTheRobotLink)
            {
                CreateStartMenuEntry(System.IO.Path.Combine(INSTALL_PATH, "Prog the robot.exe"), "Prog The Robot.lnk", "Launch Prog The Robot");
                CreateStartMenuEntry(System.IO.Path.Combine(INSTALL_PATH, fileName), "Prog The Robot installer.lnk", "Install, update and uninstall Prog The Robot");
            }

            if(needToCreateProgTheRobotDevLink)
                CreateStartMenuEntry(System.IO.Path.Combine(INSTALL_PATH, "dev", "Prog the robot.exe"), "Prog The Robot Dev.lnk", "Prog The Robot Dev Edition");

            // create the icon for the .pr files
            Stream fileImage = System.IO.File.Create(System.IO.Path.Combine(INSTALL_PATH, "FileLogo.ico"));
            fileImage.Write(Properties.Resources.FileLogo);
            fileImage.Close();

            FileAssociation fileAssociation;
            // create the file association for prog the robot dev
            if (needToCreateProgTheRobotDevLink)
            {
                fileAssociation = new FileAssociation("Prog The Robot Dev", "Prog The Robot project");
                fileAssociation.SetExtension(".pr", System.IO.Path.Combine(INSTALL_PATH, "dev", "Prog the robot.exe"), System.IO.Path.Combine(INSTALL_PATH, "FileLogo.ico"));
            }

            // create the file association for prog the robot
            if(needToCreateProgTheRobotLink)
            {
                fileAssociation = new FileAssociation("Prog The Robot", "Prog The Robot project");
                fileAssociation.SetExtension(".pr", System.IO.Path.Combine(INSTALL_PATH, "Prog the robot.exe"), System.IO.Path.Combine(INSTALL_PATH, "FileLogo.ico"));
            }

            CallBack?.Invoke();
        }

        /// <summary>
        /// Remove the program
        /// </summary>
        public string Uninstall()
        {
            if (IsFileLocked(new FileInfo(System.IO.Path.Combine(INSTALL_PATH, "Prog the robot.exe"))))
            {
                return "Please close Prog The Robot First";
            }

            RemoveDir(new string[] { PROGRAMS_PATH, $@"C:/Users/{Environment.UserName}/AppData/LocalLow/Jolan Aklin/Prog the robot" });

            FileAssociation fileAssociation = new FileAssociation("Prog The Robot");
            fileAssociation.RemoveExtension(".pr");

            fileAssociation = new FileAssociation("Prog The Robot Dev");
            fileAssociation.RemoveExtension(".pr");

            // create a batch to remove everything
            var procId = Process.GetCurrentProcess().Id;
            StreamWriter stream = System.IO.File.CreateText(UNINSTALLBAT_PATH);
            stream.WriteLine($"Taskkill /F /PID {procId}");
            stream.WriteLine($"rmdir /S /Q \"{INSTALL_PATH}\"");
            stream.WriteLine("powershell -Command \"& {Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.MessageBox]::Show('Prog The Robot was successfully removed', 'Uninstaller', 'OK', [System.Windows.Forms.MessageBoxIcon]::Information);}\"");
            stream.WriteLine($"del /Q \"{UNINSTALLBAT_PATH}\"");
            stream.Close();

            // start the batch
            Process process = new Process();
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.FileName = UNINSTALLBAT_PATH;
            process.Start();

            System.Windows.Application.Current.Shutdown();

            return null;
        }

        /// <summary>
        /// <see cref="https://www.codercream.com/check-file-use-c-code/"/>
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        private static bool IsFileLocked(FileInfo file)
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
