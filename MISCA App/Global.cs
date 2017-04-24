using System;
using System.Collections.Generic;
using System.IO;
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
        IList<long> album_id = new List<long>();

        List<CheckingWB> images = new List<CheckingWB>();
        string[] key_words = new string[10];

        double cny;
        string imagesHidden = string.Empty;
        bool isload = true, ismain = false, isImgAdded = false;
        int i = 0, rowIdx = 1, count = 1, img_count=12;
        string content = string.Empty;

        NotifyIcon nf = new NotifyIcon();
        static Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook wbook = app.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + "Products.xlsx");
        Microsoft.Office.Interop.Excel.Worksheet wsheet;

        DirectoryInfo dirInfo = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + @"\Изображения\");

        public delegate void Translation();
        public delegate void Link();
        public delegate void Images();
        public delegate void Loop();     
    }
}
