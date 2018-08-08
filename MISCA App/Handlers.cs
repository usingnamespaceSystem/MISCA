using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Diagnostics;
using System.IO;
using System.Windows.Data;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Text;
using System.Reflection;
using System.Data.OleDb;
using System.Data;

namespace MISCA_App
{
    public partial class MainWindow : Window
    {
        /// <summary>
        /// Собрать товар
        /// </summary>
        private void fusropars_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                get_curency();
            }
            catch
            {
                _cny = 8.7;
                MessageBox.Show("Курс не получен");
            }

            ;

            clear();
            WebControl_promo.Visibility = Visibility.Visible;
            TranslateText();
        }

        private void link_KeyDown(object sender, KeyEventArgs e)
        {
            if (Keyboard.IsKeyDown(Key.Enter) && Keyboard.IsKeyDown(Key.LeftShift) && size.IsFocused)
            {
                size.Text += "\n";
                size.SelectionStart = size.Text.Length;
            }
            else if (e.Key == Key.Enter)
            {
                link.Text = Clipboard.GetText();
                fusropars.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            }
            else if (e.Key == Key.LeftCtrl && e.Key == Key.S)
            {
                save.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            }
        }

        private void reload_Click(object sender, RoutedEventArgs e)
        {
            WebControl.Reload(true);
        }

        private void WebControl_LoadingFrameComplete(object sender, Awesomium.Core.FrameEventArgs e)
        {
            status.Content = "Загрузка завершена";

            if (!_isload) return;
            WebControl_promo.Source = new Uri(link.Text);
            const string script =
                @"(function() { for (var i in g_config.promotion.promoData) { return g_config.promotion.promoData[i][0].price } }())";
            WebControl_promo.LoadingFrameComplete += (obj, evt) =>
            {
                string promotion = WebControl_promo.ExecuteJavascriptWithResult(script);

                if (promotion != "undefined")
                    price.Text = promotion.Trim('"').Replace('.', ',');

                if (promotion == "undefined")
                    Findprice();

                if (price.Text.Contains("-"))
                    price.Text = price.Text.Split('-')[1].Trim();

                if (price.Text != "")
                    perc.Text = Convert.ToString(Math.Round(2500 / Convert.ToDecimal(price.Text.Replace('.', ',')), 0),
                        CultureInfo.CurrentCulture);

                if (perc.Text == "0" || perc.Text == "")
                {
                    perc.Text = "20";
                }

                WebControl_promo.Visibility = Visibility.Hidden;
            };

            _content = WebControl.ExecuteJavascriptWithResult("document.getElementsByTagName('html')[0].innerHTML");

            Findname();

            Findseller();

            Findmaterial();

            Get_images();

            _isload = false;

            auto_category();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            foreach (var file in _dirInfo.GetFiles())
            {
                file.Delete();
            }

            try
            {
                _wbook.Close(true);
            }
            catch (Exception)
            {
                MessageBox.Show("Не удалось закрыть Excel");
            }


            App.Quit();

            try
            {
                foreach (var process in Process.GetProcessesByName("EXCEL"))
                {
                    process.Kill();
                }
            }
            catch
            {
                // ignored
            }
        }

        private void img_added_Click(object sender, RoutedEventArgs e)
        {
            if (_i <= 0 || _i >= 5) return;
            _isImgAdded = true;
            img_checking_count.Content = Convert.ToInt32(img_checking_count.Content) + 1;
        }


        private void img_deleted_Click(object sender, RoutedEventArgs e)
        {
            if (_i <= 0 || _i >= 5) return;
            _isImgAdded = false;
            img_checking_count.Content = Convert.ToInt32(img_checking_count.Content) - 1;
        }

        private void forward_Click(object sender, RoutedEventArgs e)
        {
            WebControl.GoForward();
        }

        private void parse_size_Click(object sender, RoutedEventArgs e)
        {
            size_table.ItemsSource = SizeRowCollection;
            List<string> parameters = new List<string>();
            parameters.AddRange(size.Text.Split(','));
            size_table.Columns[0].Header = "Размер";
            for (int i = 1; i <= parameters.Count; i++)
            {
                size_table.Columns[i].Header = parameters[i - 1];
            }
            //size_table.Items.Add(new ListCollectionView(parameters));
            SizeRow new_row = new SizeRow() { field1 = string.Empty, field2 = string.Empty, field3 = string.Empty, field4 = string.Empty, field5 = string.Empty };
            SizeRowCollection.Add(new_row);

            _isSizeInTable = true;
        }


        private void back_Click(object sender, RoutedEventArgs e)
        {
            WebControl.GoBack();
        }

        private void change_price(object sender, TextChangedEventArgs e)
        {
            if (price.Text == string.Empty || perc.Text == string.Empty) return;
            try
            {
                final_price.Text = Math.Round(
                    (
                        (Convert.ToDouble(price.Text) * (1.07 + Convert.ToDouble(perc.Text) / 100.0) + 20) * _cny +
                        Convert.ToDouble(ship.Content)),
                    0).ToString();
                income.Content =
                    (Convert.ToDouble(final_price.Text) - Math.Round(
                         ((Convert.ToDouble(price.Text) * 1.07 + 20) * _cny + Convert.ToDouble(ship.Content)), 0))
                    .ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка при изменении цены: " + ex.Message);
            }
        }

        private void change_weight(object sender, TextChangedEventArgs e)
        {
            if (weight.Text == String.Empty) return;
            double agentsComission = 0.0;

            try
            {
                //комиссия поставщика считается либо за вес, либо единожды за посылку вцелом
                if (_agentRow.Columns[3].Value != "")
                {
                    agentsComission = (double) (Convert.ToDecimal(weight.Text) * Convert.ToDecimal(_agentRow.Columns[3].Value));
                }
                else
                {
                    agentsComission = (double) (Convert.ToDecimal(weight.Text) * Convert.ToDecimal(_agentRow.Columns[4].Value));
                }

                //комиссия поставщика считается либо за вес, либо единожды за посылку вцелом
                double rus_ship = (double)(Convert.ToDecimal(weight.Text) * Convert.ToDecimal(_agentRow.Columns[6].Value));
                ship.Content =  ((double)(Convert.ToDecimal(weight.Text) * Convert.ToDecimal(_agentRow.Columns[5].Value)) + agentsComission + rus_ship).ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при изменении веса: {ex.Message}");
            }
        }


        private void category_DropDownClosed(object sender, System.EventArgs e)
        {
            foreach (Microsoft.Office.Interop.Excel.Worksheet sh in _wbook.Worksheets)
            {
                if (sh.Name != category.SelectionBoxItem.ToString())
                {
                    continue;
                }

                weight.Text = sh.Cells[2, 5].Value.ToString();
                double agentsComission = 0.0;

                //комиссия поставщика считается либо за вес, либо единожды за посылку вцелом
                if (_agentRow.Columns[3].Value != "")
                {
                    agentsComission = (double) (Convert.ToDecimal(weight.Text) * Convert.ToDecimal(_agentRow.Columns[3].Value));
                }
                else
                {
                    agentsComission = (double) Convert.ToDecimal(_agentRow.Columns[4].Value);
                }

                //комиссия поставщика считается либо за вес, либо единожды за посылку вцелом
                double rus_ship = (double)(Convert.ToDecimal(weight.Text) * Convert.ToDecimal(_agentRow.Columns[6].Value));
                ship.Content = ((double)(Convert.ToDecimal(weight.Text) * Convert.ToDecimal(_agentRow.Columns[5].Value)) + agentsComission + rus_ship).ToString();
                return;
            }
        }

        private void agent_DropDownClosed(object sender, System.EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Worksheet wsheetAgents = _wbookAgents.Worksheets[1];
            foreach (Microsoft.Office.Interop.Excel.Range row in wsheetAgents.UsedRange.Rows)
            {
                if (row.Columns[6].Text == agent.SelectionBoxItem.ToString())
                {
                    _agentRow = row;
                }
            }
        }


        private void WebControl_LoadingFrame(object sender, Awesomium.Core.LoadingFrameEventArgs e)
        {
            status.Content = "Загрузка страницы...";
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Auth();
        }

        private void price_KeyDown(object sender, KeyEventArgs e)
        {
            if (!((e.Key.GetHashCode() >= 34) && (e.Key.GetHashCode() <= 43)) &&
                !((e.Key.GetHashCode() >= 74) && (e.Key.GetHashCode() <= 83)) && e.Key.GetHashCode() == 73)
            {
                e.Handled = true;
            }
        }

        private void grid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            string fileName = null, sheet = null;
            string type = (sender as DataGrid).Tag.ToString();
            DataGrid dg = sender as DataGrid;
            DataRowView row = (DataRowView)dg.SelectedItems[0];

            switch (type)
            {
                case "orders":
                    {
                        fileName = order_file;
                        sheet = orders_sheet;
                        break;
                    }
                case "stock":
                    {
                        fileName = product_file;
                        sheet = _category_for_stock;
                        break;
                    }
            }

            string col = e.Column.Header.ToString();
            string value = ((TextBox)e.EditingElement).Text;
            string id = row["Артикул"].ToString();
            string connectionString = string.Format("provider=Microsoft.Jet.OLEDB.4.0; data source={0};Extended Properties=Excel 8.0;", fileName);

            if (type == "stock")
                _arts_edited.Add(id);

            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.Parameters.AddWithValue("@val", value);
                cmd.Parameters.AddWithValue("@art", id);
                cmd.CommandText = string.Format("UPDATE [{0}$] SET {1}=@val WHERE [Артикул]=@art", sheet, col);
                cmd.ExecuteNonQuery();
            }
        }


        private int find_excel_col(Microsoft.Office.Interop.Excel.Worksheet fwsheet, string fheader)
        {
            foreach (Microsoft.Office.Interop.Excel.Range cell in fwsheet.Rows[1])
            {
                if (cell.Value == fheader)
                {
                    return cell.Column;
                }
            }

            return 0;
        }

        private int find_excel_row(Microsoft.Office.Interop.Excel.Worksheet fwsheet, string fheader, string fvalue)
        {
            int fcolumn = 0;

            fcolumn = find_excel_col(fwsheet, fheader);

            if (fcolumn == 0)
                return 0;

            foreach (Microsoft.Office.Interop.Excel.Range row in fwsheet.UsedRange.Rows)
            {
                if (row.Columns[fcolumn].Value == fvalue)
                {
                    return row.Row;
                }
            }

            return 0;
        }

        private void upload_to_inst(object sender, RoutedEventArgs e) { }

        private void orders_grid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            e.Column.MaxWidth = 200;
        }

        private void inst_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            try
            {
                DataGrid dg = sender as DataGrid;
                DataRowView row = (DataRowView)dg.SelectedItems[0];
                WebControl_inst.Source = new Uri(row["Главн изобр"].ToString());
                inst_caption.Text = row["Наименование"].ToString() + "\n" + row["Стоимость(р)"].ToString();
            }
            catch { }
        }

        private void inst_CatChanged(object sender, RoutedEventArgs e)
        {
            string connectionString = string.Format("provider=Microsoft.ACE.OLEDB.12.0; data source={0};Extended Properties=Excel 8.0;", product_file);
            DataSet data = new DataSet();

            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                var dataTable = new DataTable();
                string query = string.Format("SELECT * FROM [{0}$]", ((RadioButton)sender).Content.ToString());
                OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
                adapter.Fill(dataTable);
                data.Tables.Add(dataTable);
            }
            product_grid.ItemsSource = data.Tables["Table1"].DefaultView;
        }


        // при изменении даты необходимо обновлять статистику
        private void date_stat_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (date_start.SelectedDate == null || date_end.SelectedDate == null)
                return;

            string connectionString = string.Format("provider=Microsoft.ACE.OLEDB.12.0; data source={0};Extended Properties=Excel 8.0;", order_file);
            DataSet ds_orders = new DataSet();
            DataSet ds_cats = new DataSet();

            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                con.Open();
                var stat_regions = new DataTable();
                var stat_cats = new DataTable();
                float sum = 0, count = 0;

                //cmd.Parameters.AddWithValue("@start", date_start.SelectedDate.Value);
                //cmd.Parameters.AddWithValue("@end", date_end.SelectedDate.Value);
  
                //string query_region = "SELECT [Регион], COUNT([Регион]) as [Количество заказов] FROM [2017$] WHERE [Дата] >= \'@start\' AND [Дата] <= \'@end\' GROUP BY [Регион]";
                string query_region = String.Format("SELECT [Регион], COUNT([Регион]) as [Количество заказов] FROM [{0}$] WHERE [Регион] IS NOT NULL GROUP BY [Регион] ORDER BY COUNT([Регион]) DESC", orders_sheet);
                OleDbDataAdapter adapter_region = new OleDbDataAdapter(query_region, con);
                adapter_region.Fill(stat_regions);
                ds_orders.Tables.Add(stat_regions);

                string query_category = String.Format("SELECT [Товарная категория], COUNT([Товарная категория]) as [Количество заказов]" +
                    "FROM [{0}$] WHERE [Товарная категория] IS NOT NULL GROUP BY [Товарная категория] ORDER BY COUNT([Товарная категория]) DESC", orders_sheet);
                OleDbDataAdapter adapter_category = new OleDbDataAdapter(query_category, con);
                adapter_category.Fill(stat_cats);
                ds_cats.Tables.Add(stat_cats);

                OleDbCommand cmd = new OleDbCommand(String.Format("SELECT Sum([Прибыль]) FROM [{0}$]", orders_sheet), con);
                OleDbDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                    stat_income.Content = reader[0].ToString();

                cmd = new OleDbCommand(String.Format("SELECT COUNT([Источник]) FROM [{0}$] WHERE [Источник]=\"ВК\" ", orders_sheet), con);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                    stat_vk.Content = reader[0].ToString();

                cmd = new OleDbCommand(String.Format("SELECT COUNT([Источник]) FROM [{0}$] WHERE [Источник]=\"Инстаграм\" ", orders_sheet), con);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                    stat_inst.Content = reader[0].ToString();

                cmd = new OleDbCommand(String.Format("SELECT [Стоимость] FROM [{0}$] WHERE [№] IS NOT NULL", orders_sheet), con);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    try
                    {
                        sum += (float) Convert.ToDecimal(reader[0]);
                        count++;
                    }
                    catch
                    { continue; }
                }
                stat_summ.Content = (sum/count).ToString();
            }
            region_grid.ItemsSource = ds_orders.Tables["Table1"].DefaultView;
            cat_grid.ItemsSource = ds_cats.Tables["Table1"].DefaultView;
        }

        private void check_available_click(object sender, RoutedEventArgs e)
        {
            _arts_edited.Clear();

            string reply = string.Empty;
            string fileName = AppDomain.CurrentDomain.BaseDirectory + "Products.xlsx"; ;
            int art_column = 0;
            List<string> arts = new List<string>();

            for (int i = 1; i <= LastColumn; i++)
            {
                string header = _wbook.Worksheets[1].Cells[1, i].Text;
                if (header == "Ссылка")
                {
                    _linkColumn = i;
                }
                if (header == "Артикул")
                {
                    art_column = i;
                }
            }

            foreach (RadioButton ch in category_panel.Children)
            {
                Microsoft.Office.Interop.Excel.Worksheet wsheet = _wbook.Worksheets[ch.Content];
                if (ch.IsChecked == true)
                {
                    foreach (Microsoft.Office.Interop.Excel.Range row in wsheet.UsedRange.Rows)
                    {
                        if (row.Row == 1)
                            continue;

                        using (WebClient client = new WebClient())
                        {
                            try
                            {
                                client.Encoding = System.Text.Encoding.GetEncoding("GB2312");
                                reply = client.DownloadString(row.Columns[_linkColumn].Text);
                            }
                            catch (WebException ex)
                            {
                                if (ex.Status == WebExceptionStatus.ProtocolError && ex.Response != null)
                                    arts.Add(row.Columns[art_column].Text);
                            }
                        }

                        if (reply.Contains("此宝贝已下架") || reply.Contains("此商品已下架") || reply.Contains("您查看的宝贝不存在"))
                        {
                            arts.Add(row.Columns[art_column].Text);
                            reply = string.Empty;
                        }
                    }
                }

                string connectionString = string.Format("provider=Microsoft.ACE.OLEDB.12.0; data source={0};Extended Properties=Excel 8.0;", fileName);
                DataSet data = new DataSet();
                _category_for_stock = ch.Content.ToString();

                using (OleDbConnection con = new OleDbConnection(connectionString))
                {
                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    var dataTable = new DataTable();
                    string str = String.Join(", ", arts);
                    //cmd.Parameters.AddWithValue("@arts", str);
                    //string query = string.Format("SELECT * FROM [{0}$] WHERE [Артикул] in (@arts)", ch.Content);
                    string query = string.Format("SELECT * FROM [{0}$] WHERE [Артикул] in ({1})", ch.Content, str);
                    OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
                     adapter.Fill(dataTable);
                    data.Tables.Add(dataTable);
                }
                stock_grid.ItemsSource = data.Tables["Table1"].DefaultView;
                break;
            }
            
        }

        private void inst_Click(object sender, RoutedEventArgs e)
        {
            string user = Config.Read("username", "instagram");
            string pass = Config.Read("password", "instagram");
            WebClient wc = new WebClient();
            string img_path = images_file + "\\inst.jpg";
            wc.DownloadFile(WebControl_inst.Source, img_path);
            InstargamUpload.UploadImage(user, pass, img_path, inst_caption.Text);
        }

    }
}