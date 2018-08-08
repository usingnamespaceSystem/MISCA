using System.Windows;
using System.Windows.Controls;
using System.Collections.ObjectModel;
using System.Net;
using System;
using System.Data;
using System.Collections.Generic;
using System.Data.OleDb;

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
            string connectionString = string.Format("provider=Microsoft.ACE.OLEDB.12.0; data source={0};Extended Properties=Excel 8.0;", order_file);
            orders_grid.ItemsSource = Parse(order_file).Tables["Table1"].DefaultView;

            connectionString = string.Format("provider=Microsoft.ACE.OLEDB.12.0; data source={0};Extended Properties=Excel 8.0;", product_file);
            product_grid.ItemsSource = Parse(product_file).Tables["Table1"].DefaultView;
            //foreach (DataGridColumn col in orders_grid.Columns)
            //    col.MaxWidth = new DataGridLenght(200);

            //int ch_top = 0, ch_left= 0, margin_top = 20;
            //заполняем список категорий из excel-файла Products
            foreach (Microsoft.Office.Interop.Excel.Worksheet sh in _wbook.Worksheets)
            {
                if (!sh.Name.Contains("nul"))
                {
                    category.Items.Add(sh.Name);

                    RadioButton cat_select = new RadioButton();
                    cat_select.GroupName = "category_for_check_stock";
                    cat_select.Content = (string)sh.Name;
                    category_panel.Children.Add(cat_select);

                    RadioButton cat_inst_select = new RadioButton();
                    cat_inst_select.GroupName = "category_for_inst";
                    cat_inst_select.Content = (string)sh.Name;
                    cat_inst_select.Checked += inst_CatChanged;
                    category_inst_panel.Children.Add(cat_inst_select);
                }
            }

            //заполняем список посредников из excel-файла Agents
            Microsoft.Office.Interop.Excel.Worksheet wsheet_agents = _wbookAgents.Worksheets[1];
            foreach (Microsoft.Office.Interop.Excel.Range row in wsheet_agents.UsedRange.Rows)
            {
                if (row.Row == 1)
                    continue;

                agent.Items.Add(row.Columns[1].Text);
                if (row.Columns[7].Text == "да")
                {
                    _agentRow = row;
                }
            }
            agent.SelectedItem = _agentRow.Columns[1].Text;
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

        ////для редактирования строк в гриде наличия
        //public class StockRow
        //{
        //    public string stock_category { get; set; }
        //    public string stock_article { get; set; }
        //    public string stock_status { get; set; }
        //    public string stock_name { get; set; }
        //    public string stock_link { get; set; }
        //    public string stock_seller { get; set; }
        //    public string stock_material { get; set; }
        //    public string stock_size { get; set; }
        //    public string stock_price { get; set; }
        //    public string stock_percent { get; set; }
        //    public string stock_shipping { get; set; }
        //    public string stock_summary { get; set; }
        //}
        //private ObservableCollection<StockRow> _stock_rows;
        //public ObservableCollection<StockRow> StockRowCollection
        //{
        //    get
        //    {
        //        if (_stock_rows == null)
        //            _stock_rows = new ObservableCollection<StockRow>();
        //        return _stock_rows;
        //    }
        //}
    }
}
