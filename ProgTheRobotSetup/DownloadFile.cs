using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ProgTheRobotSetup
{
    public class DownloadFile
    {
        /// <summary>
        /// Fired when the download process changes
        /// </summary>
        public EventHandler<DownloadProgressChangedEventArgs> onProgressChanged;
        /// <summary>
        /// Fired when the download is finished
        /// </summary>
        public EventHandler<EventArgs> onDownloadFinished;

        /// <summary>
        /// Download file async
        /// </summary>
        /// <param name="ressource">Where the file is stored</param>
        /// <param name="savePath">The path where the file will be saved to</param>
        /// <param name="endCallback">Called when the download is finished</param>
        public async void DownloadFromUrl(Uri ressource, string savePath, Action endCallback)
        {
            WebClient webClient = new WebClient();
            webClient.DownloadProgressChanged += ProgressChanged;
            await webClient.DownloadFileTaskAsync(ressource, savePath);
            onDownloadFinished?.Invoke(this, EventArgs.Empty);
            endCallback?.Invoke();
        }

        private void ProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            onProgressChanged?.Invoke(sender, e);
        }
    }
}
