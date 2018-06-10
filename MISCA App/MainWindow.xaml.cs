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
            string fileName = AppDomain.CurrentDomain.BaseDirectory + "заказы.xlsx";
            string connectionString = string.Format("provider=Microsoft.ACE.OLEDB.12.0; data source={0};Extended Properties=Excel 8.0;", fileName);
            DataSet ds = Parse(fileName);
            orders_grid.ItemsSource = Parse(AppDomain.CurrentDomain.BaseDirectory + "заказы.xlsx").Tables["Table1"].DefaultView;
            //foreach (DataGridColumn col in orders_grid.Columns)
            //    col.MaxWidth = new DataGridLenght(200);

            //int ch_top = 0, ch_left= 0, margin_top = 20;
            //заполняем список категорий из excel-файла Products
            foreach (Microsoft.Office.Interop.Excel.Worksheet sh in wbook.Worksheets)
            {
                if (!sh.Name.Contains("nul"))
                {
                    category.Items.Add(sh.Name);
                    RadioButton cat_select = new RadioButton();
                    cat_select.GroupName = "category_for_check_stock";
                    //cat_check.Margin = new Thickness(ch_left, ch_top, 0, 0);
                    cat_select.Content = (string)sh.Name;
                    category_panel.Children.Add(cat_select);
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

    }
}
