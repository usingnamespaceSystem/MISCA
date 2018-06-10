using System;
using System.Security;
using InstaSharp;
using InstaSharp.Models;
using MISCA_App.InstagramAPI;
using MISCA_App.InstagramAPI.Models;

namespace MISCA_App
{
    public class InstargamUpload
    {
        private void UploadImage(string username, string password, string imagePath, string caption)
        {
            var uploader = new InstagramUploader(username, ConvertToSecureString(password));
            uploader.InvalidLoginEvent += InvalidLoginEvent;
            uploader.ErrorEvent += ErrorEvent;
            uploader.OnCompleteEvent += OnCompleteEvent;
            uploader.OnLoginEvent += OnLoginEvent;
            uploader.SuccessfulLoginEvent += SuccessfulLoginEvent;
            uploader.OnMediaConfigureStarted += OnMediaConfigureStarted;
            uploader.OnMediaUploadStartedEvent += OnMediaUploadStartedEvent;
            uploader.OnMediaUploadeComplete += OnmediaUploadCompleteEvent;
            uploader.UploadImage(imagePath, caption);
        }

        private static SecureString ConvertToSecureString(string strPassword)
        {
            var secureStr = new SecureString();
            if (strPassword.Length <= 0) return secureStr;
            foreach (var c in strPassword.ToCharArray()) 
                secureStr.AppendChar(c);

            return secureStr;
        }

        private static void OnMediaUploadStartedEvent(object sender, EventArgs e)
        {

        }

        private static void OnmediaUploadCompleteEvent(object sender, EventArgs e)
        {

        }


        private static void OnMediaConfigureStarted(object sender, EventArgs e)
        {

        }

        private static void SuccessfulLoginEvent(object sender, EventArgs e)
        {
        }

        private static void OnLoginEvent(object sender, EventArgs e)
        {
        }

        private static void OnCompleteEvent(object sender, EventArgs e)
        {
            foreach (var image in ((UploadResponse) e).Images)
            {
               
            }
        }

        private static void ErrorEvent(object sender, EventArgs e)
        {
        }

        private static void InvalidLoginEvent(object sender, EventArgs e)
        {
        }
    }
}