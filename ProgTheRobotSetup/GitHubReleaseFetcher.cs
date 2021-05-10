using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Octokit;

namespace ProgTheRobotSetup
{
    public static class GitHubReleaseFetcher
    {

        public enum DownloadableFiles
        {
            ProgTheRobot,
            SoundPack,
            DemoPack,
            ProgTheRobotDev,
        }

        private static Dictionary<DownloadableFiles, string> downloadableFiles = new Dictionary<DownloadableFiles, string>() {
            {DownloadableFiles.ProgTheRobot, "ProgTheRobot.zip" },
            {DownloadableFiles.SoundPack, "SoundPack.zip" },
            {DownloadableFiles.DemoPack, "DemoPack.zip" },
            {DownloadableFiles.ProgTheRobotDev, "ProgTheRobotDev.zip" },
        };

        private const string REPO_NAME = "progtherobot";
        private const string REPO_OWNER = "jolanaklin";
        private const string CLIENT_HEADER_INFORMATIONS = "ProgTheRobotSetup";

        private static GitHubClient gitHubClient;
        private static Release latestRelease;
        public static Release LatestRelease { get => latestRelease; private set => latestRelease = value; }


        /// <summary>
        /// Create the GitHubClient
        /// </summary>
        /// <returns>The newly created GitHubClient</returns>
        public static GitHubClient CreateGitHubClient()
        {
            try
            {
                gitHubClient = new GitHubClient(new ProductHeaderValue(CLIENT_HEADER_INFORMATIONS));
                return gitHubClient;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Find the latest release
        /// </summary>
        public static async void FindLatestRelease()
        {
            try
            {
                if (gitHubClient == null)
                {
                    if (CreateGitHubClient() == null)
                        return;
                }
                IReadOnlyList<Release> releases = await gitHubClient.Repository.Release.GetAll(REPO_OWNER, REPO_NAME);
                latestRelease = releases[0];
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Find the latestRelease
        /// </summary>
        /// <param name="callBack">Will be called when the latest release is retrieded</param>
        public static void FindLatestRelease(Action callBack)
        {
            FindLatestRelease();
            callBack?.Invoke();
        }

        /// <summary>
        /// Get the download url of the latest release asset
        /// </summary>
        /// <param name="asset">The asset to get</param>
        /// <returns>The url to the specified asset</returns>
        public static string GetReleaseAssetUrl(DownloadableFiles asset)
        {
            string filename = downloadableFiles[asset];
            return Array.Find<ReleaseAsset>(GitHubReleaseFetcher.LatestRelease.Assets.ToArray(), x => x.Name == filename).BrowserDownloadUrl;
        }
    }
}
