using System;

namespace MISCA_App.InstagramAPI.Models
{
    public class NormalResponse : EventArgs
    {
        public string Status { get; set; }
        public string Message { get; set; }
    }
}