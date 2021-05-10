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
        /// Uninstall the app
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Uninstall(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Do you really want to uninstall this program ?", "Uninstall", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                Install install = new Install(PROGRAMS_PATH, INSTALL_PATH, TEMP_PATH, DL_FILE_NAME, UNINSTALLBAT_PATH);
                string info = install.Uninstall();
                if (info != null)
                    MessageBox.Show(info, "Uninstall", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void InstallSettingsNext(object sender, RoutedEventArgs e)
        {
            ShowNextGrid();
            StartInstall();
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

            if(GitHubReleaseFetcher.CreateGitHubClient() == null)
            {
                MessageBox.Show("Impossible de se connecter à GitHub" + Environment.NewLine + "Verifier que votre PC est connecté à Internet", "Connecting to GitHub", MessageBoxButton.OK, MessageBoxImage.Error); ShowMain();
            }
            else
            {
                GitHubReleaseFetcher.FindLatestRelease(() => { button.IsEnabled = true; });
            }

        }

        private void StartInstall()
        {
            Install install = new Install(PROGRAMS_PATH, INSTALL_PATH, TEMP_PATH, DL_FILE_NAME, UNINSTALLBAT_PATH);
            install.PreInstall();

            string url = GitHubReleaseFetcher.GetReleaseAssetUrl(GitHubReleaseFetcher.DownloadableFiles.ProgTheRobot);

            DownloadFile dlFile = new DownloadFile();
            dlFile.onProgressChanged += (object sender, DownloadProgressChangedEventArgs e) => { DownLoadProgress.Value = e.ProgressPercentage; };
            dlFile.DownloadFromUrl(new Uri(url), System.IO.Path.Combine(TEMP_PATH, DL_FILE_NAME),
            () => {
                install.InstallApp(() => { ShowNextGrid(); });
            });
        }

        /// <summary>
        /// Show the main grid
        /// </summary>
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
        /// Show the release tag
        /// </summary>
        private void ShowReleaseTag()
        {
            SolidColorBrush solidColor = null;
            if (GitHubReleaseFetcher.LatestRelease.Prerelease)
            {
                solidColor = new SolidColorBrush(Color.FromRgb(Convert.ToByte(255), Convert.ToByte(133), Convert.ToByte(0)));
            }
            else
            {
                solidColor = new SolidColorBrush(Color.FromRgb(Convert.ToByte(0), Convert.ToByte(128), Convert.ToByte(0)));
            }
            VerLabel.Background = solidColor;
        }
    }
}
