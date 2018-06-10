using System;

namespace MISCA_App.InstagramAPI.Models
{
    public class ErrorResponse : EventArgs
    {
        public string Status { get; set; }
        public string Message { get; set; }
    }
}