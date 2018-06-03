using System;
using System.Collections.Generic;

namespace MISCA_App.InstagramAPI.Models
{
    public class UploadResponse : EventArgs
    {
        public List<InstagramMedia> Images { get; set; }

        public class InstagramMedia
        {
            public string Url { get; set; }
            public int Width { get; set; }
            public int Height { get; set; }
        }
    }
}