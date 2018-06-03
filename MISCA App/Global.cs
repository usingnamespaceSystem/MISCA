﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Forms;
using VkNet;

namespace MISCA_App
{
    public partial class MainWindow : Window
    {
        private static readonly IniFile Config = new IniFile();
        [DllImport("Kernel32")]
        public static extern void AllocConsole();

        [DllImport("Kernel32")]
        public static extern void FreeConsole();

        VkApi vk = new VkApi();
        long[] extraPhotos = new long[4];
        List<long> album_id = new List<long>();
        List<string> list_of_categories = new List<string>();
        List<CheckingWB> images = new List<CheckingWB>();
        string[] key_words = new string[10];

        double cny;
        string imagesHidden = string.Empty;
        bool isload = true, ismain = false, isImgAdded = false, isSizeInTable = false;
        int i = 0, rowIdx = 1, count = 1, img_count=12, link_column = 0, last_column = 11;
        string content = string.Empty;
        int cat_id = 1;

        NotifyIcon nf = new NotifyIcon();
        static Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

        private static readonly string WbookPath = Config.Read("products", "app");
        private static readonly string WbookAgentsPath = Config.Read("agents", "app");
        private static readonly string ImagesPath = Config.Read("images", "app");
        
        Microsoft.Office.Interop.Excel.Workbook wbook = app.Workbooks.Open(WbookPath);
        Microsoft.Office.Interop.Excel.Workbook wbook_agents = app.Workbooks.Open(WbookAgentsPath);
        
        Microsoft.Office.Interop.Excel.Worksheet wsheet;
        Microsoft.Office.Interop.Excel.Range agent_row;

        DirectoryInfo dirInfo = new DirectoryInfo(ImagesPath);

        public delegate void Translation();
        public delegate void Link();
        public delegate void Images();
        public delegate void Loop();

        
    }
}
