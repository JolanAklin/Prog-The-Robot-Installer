using System;
using System.Collections.Generic;
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

            string path = $@"C:\Users\{Environment.UserName}\Desktop\Prog_the_robot_v1-0a.zip";
            WebClient webClient = new WebClient();
            webClient.DownloadProgressChanged += onProgressChanged;
            webClient.DownloadFileAsync(new Uri("https://github.com/JolanAklin/ProgTheRobot/releases/download/v1.0a/Prog_the_robot_v1-0a.zip"), path);
        }

        private void onProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            DownLoadProgress.Value = e.ProgressPercentage;
        }
    }
}
