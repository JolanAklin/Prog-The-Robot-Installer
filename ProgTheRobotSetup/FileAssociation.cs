using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProgTheRobotSetup
{
    /// <summary>
    /// <see cref="https://social.msdn.microsoft.com/Forums/en-US/6e61be53-86cd-4761-bcd9-34a4f5a75503/how-i-can-create-a-file-association"/>
    /// </summary>

    class FileAssociation
    {
        [System.Runtime.InteropServices.DllImport("Shell32.dll")]
        private static extern int SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);
        public FileAssociation(string applicationName)
        {
            this.applicationName = applicationName;
            this.fileTypeName = applicationName + " file";
        }

        public FileAssociation(string applicationName, string fileTypeName)
        {
            this.applicationName = applicationName;
            this.fileTypeName = fileTypeName;
        }

        private string fileTypeName;

        public string FileTypeName
        {
            get { return fileTypeName; }
            set { fileTypeName = value; }
        }

        private string applicationName;

        public string ApplicationName
        {
            get { return applicationName; }
            set { applicationName = value; }
        }

        public void SetExtension(string extension, string appPath)
        {
            RegistryKey root = Registry.ClassesRoot;
            RegistryKey key;

            RegistryKey exist = root.OpenSubKey(extension);
            if(exist != null)
                root.DeleteSubKeyTree(extension);
            //Create extension key that refers to application name
            key = root.CreateSubKey(extension);
            key.SetValue("", applicationName, RegistryValueKind.String);
            key.Close();

            exist = root.OpenSubKey(applicationName);
            if (exist != null)
                root.DeleteSubKeyTree(applicationName);
            //Create application key that is referenced by extension key
            RegistryKey programKey = root.CreateSubKey(applicationName);
            programKey.SetValue("", fileTypeName, RegistryValueKind.String);

            // set the file icon
            key = programKey.CreateSubKey("DefaultIcon");
            key.SetValue("", appPath, RegistryValueKind.String);
            key.Close();

            //Create open command for application type (enables multiple files to use same application key)
            key = programKey.CreateSubKey(@"shell\open\command");
            key.SetValue("", appPath + " %1", RegistryValueKind.String);
            key.Close();
            programKey.Close();
            root.Close();

            // tell the shell to remake the icon cache
            SHChangeNotify(0x08000000, 0,IntPtr.Zero, IntPtr.Zero);
        }

        public bool Verify(string extension)
        {
            RegistryKey root = Registry.ClassesRoot;
            RegistryKey key;

            key = root.OpenSubKey(extension);

            //If extension key does not exist, return false
            if (key == null)
            {
                return false;
            }

            string appName = (string)key.GetValue("");

            //If configured application name does not match expected, return false
            if (appName != applicationName)
            {
                return false;
            }

            //Assume that application references this one
            return true;
        }
    }
}