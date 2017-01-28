using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using VkNet;

namespace MISCA_App
{
    public partial class MainWindow : Window
    {
        [DllImport("Kernel32")]
        public static extern void AllocConsole();

        [DllImport("Kernel32")]
        public static extern void FreeConsole();


        VkApi vk = new VkApi();
        long[] extraPhotos = new long[4];
        int count = 1;


        NotifyIcon nf = new NotifyIcon();
        int i = 1, rowIdx = 1;
        static Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "Products.xlsx");
        Microsoft.Office.Interop.Excel.Worksheet ws;

        public delegate void Translation();
        public delegate void Link();
        public delegate void Images();
        public delegate void Loop();
        string content = string.Empty;
        List<CheckingWB> images = new List<CheckingWB>();
        float usd;
        bool load = true;
        string imagesHidden = string.Empty;

    }
}
