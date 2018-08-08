using System;
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

        private readonly VkApi _vk = new VkApi();
        private readonly long[] _extraPhotos = new long[4];
        private readonly List<long> _albumId = new List<long>();
        private readonly List<CheckingWB> _images = new List<CheckingWB>();
        private string[] _keyWords = new string[10];
        private  List<string> _arts_edited = new List<string>();

        private double _cny;
        private string _imagesHidden = string.Empty,
            _category_for_stock = string.Empty;

        private bool _isload = true,
            _ismain,
            _isImgAdded,
            _isSizeInTable;

        private int _i = 0,
            _rowIdx = 1,
            _count = 1,
            _imgCount = 12,
            _linkColumn = 0,
            LastColumn = 11;

        private string _content = string.Empty;
        private int _catId = 1;

        private readonly NotifyIcon _nf = new NotifyIcon();
        static readonly Microsoft.Office.Interop.Excel.Application App;

        private static readonly string product_file = Config.Read("products", "app");
        private static readonly string agent_file = Config.Read("agents", "app");
        private static readonly string images_file = Config.Read("images", "app");
        private static readonly string order_file = Config.Read("orders", "app");
        private static readonly string orders_sheet = Config.Read("orders_sheet", "app");

        private readonly Microsoft.Office.Interop.Excel.Workbook _wbook = App.Workbooks.Open(product_file);
        private readonly Microsoft.Office.Interop.Excel.Workbook _wbookAgents = App.Workbooks.Open(agent_file);

        private Microsoft.Office.Interop.Excel.Worksheet _wsheet;
        private Microsoft.Office.Interop.Excel.Range _agentRow;

        private readonly DirectoryInfo _dirInfo = new DirectoryInfo(images_file);

        static MainWindow()
        {
            App = new Microsoft.Office.Interop.Excel.Application();
        }
    }
}