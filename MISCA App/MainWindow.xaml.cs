using System.Windows;
using System.Windows.Controls;
using System.Collections.ObjectModel;
using System.Net;
using System;
using System.Collections.Generic;

namespace MISCA_App
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            //int ch_top = 0, ch_left= 0, margin_top = 20;
            //заполняем список категорий из excel-файла Products
            foreach (Microsoft.Office.Interop.Excel.Worksheet sh in wbook.Worksheets)
            {
                if (!sh.Name.Contains("nul"))
                {
                    category.Items.Add(sh.Name);
                    CheckBox cat_check = new CheckBox();
                    //cat_check.Margin = new Thickness(ch_left, ch_top, 0, 0);
                    cat_check.Content = sh.Name;
                    category_panel.Children.Add(cat_check);
                    //ch_top += margin_top;
                }
            }

            //заполняем список посредников из excel-файла Agents
            Microsoft.Office.Interop.Excel.Worksheet wsheet_agents = wbook_agents.Worksheets[1];
            foreach (Microsoft.Office.Interop.Excel.Range row in wsheet_agents.UsedRange.Rows)
            {
                if (row.Row == 1)
                    continue;

                agent.Items.Add(row.Columns[1].Text);
                if (row.Columns[6].Text == "да")
                {
                    agent_row = row;
                }
            }
            agent.SelectedItem = agent_row.Columns[1].Text;
        }
        //для редактирования строк в гриде размеров
        public class SizeRow
        {
            public string field1 { get; set; }
            public string field2 { get; set; }
            public string field3 { get; set; }
            public string field4 { get; set; }
            public string field5 { get; set; }
        }
        private ObservableCollection<SizeRow> _size_rows;
        public ObservableCollection<SizeRow> SizeRowCollection
        {
            get
            {
                if (_size_rows == null)
                    _size_rows = new ObservableCollection<SizeRow>();
                return _size_rows;
            }
        }

        //для редактирования строк в гриде наличия
        public class StockRow
        {
            public string stock_article { get; set; }
            public string stock_status { get; set; }
            public string stock_name { get; set; }
            public string stock_link { get; set; }
            public string stock_seller { get; set; }
            public string stock_material { get; set; }
            public string stock_size { get; set; }
            public string stock_price { get; set; }
            public string stock_percent { get; set; }
            public string stock_shipping { get; set; }
            public string stock_summary { get; set; }
        }
        private ObservableCollection<StockRow> _stock_rows;
        public ObservableCollection<StockRow> StockRowCollection
        {
            get
            {
                if (_stock_rows == null)
                    _stock_rows = new ObservableCollection<StockRow>();
                return _stock_rows;
            }
        }

    }
}
