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
            var parameters = new List<string>();
            parameters.AddRange(size.Text.Split(','));
            for (int i = 1; i <= parameters.Count; i++)
            {
                size_table.Columns[i - 1].Header = parameters[i - 1];
            }

            //size_table.Items.Add(new ListCollectionView(parameters));
            var newRow = new SizeRow()
            {
                field1 = string.Empty,
                field2 = string.Empty,
                field3 = string.Empty,
                field4 = string.Empty,
                field5 = string.Empty
            };
            SizeRowCollection.Add(newRow);
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
            var agentsComission = 0;
            try
            {
                //комиссия поставщика считается либо за вес, либо единожды за посылку вцелом
                if (_agentRow.Columns[3].Value)
                {
                    agentsComission = weight.Text * _agentRow.Columns[3].Value;
                }
                else
                {
                    agentsComission = weight.Text * _agentRow.Columns[4].Value;
                }

                //комиссия поставщика считается либо за вес, либо единожды за посылку вцелом
                ship.Content = weight.Text * _agentRow.Columns[5].Value + agentsComission;
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

                weight.Text = Convert.ToInt32((sh.Cells[2, 5] as Microsoft.Office.Interop.Excel.Range)?.Value);
                var agentsComission = 0;
                //комиссия поставщика считается либо за вес, либо единожды за посылку вцелом
                if (_agentRow.Columns[3].Value)
                {
                    agentsComission = weight.Text * _agentRow.Columns[3].Value;
                }
                else
                {
                    agentsComission = weight.Text * _agentRow.Columns[4].Value;
                }

                //комиссия поставщика считается либо за вес, либо единожды за посылку вцелом
                ship.Content = weight.Text * _agentRow.Columns[5].Value + agentsComission;
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
                        fileName = AppDomain.CurrentDomain.BaseDirectory + "заказы.xlsx";
                        sheet = "2017";
                        break;
                    }
                case "stock":
                    {
                        fileName = AppDomain.CurrentDomain.BaseDirectory + "Products.xlsx";
                        sheet = category_for_stock;
                        break;
                    }
            }

            string col = e.Column.Header.ToString();
            string value = ((TextBox)e.EditingElement).Text;
            string id = row["Артикул"].ToString();
            string connectionString = string.Format("provider=Microsoft.Jet.OLEDB.4.0; data source={0};Extended Properties=Excel 8.0;", fileName);

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

        // при изменении даты необходимо обновлять статистику
        private void date_stat_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (date_start.SelectedDate == null || date_end.SelectedDate == null)
                return;

            string fileName = AppDomain.CurrentDomain.BaseDirectory + "заказы.xlsx"; ;
            string connectionString = string.Format("provider=Microsoft.ACE.OLEDB.12.0; data source={0};Extended Properties=Excel 8.0;", fileName);
            DataSet data = new DataSet();

            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                var stat_regions = new DataTable();
                var dataTable = new DataTable();
                float sum = 0, count = 0;

                cmd.Parameters.AddWithValue("@start", date_start.SelectedDate.Value);
                cmd.Parameters.AddWithValue("@end", date_end.SelectedDate.Value);

                //string query_region = "SELECT [Регион], COUNT([Регион]) as [Количество заказов] FROM [2017$] WHERE [Дата] >= \'@start\' AND [Дата] <= \'@end\' GROUP BY [Регион]";
                string query_region = "SELECT [Регион], COUNT([Регион]) as [Количество заказов] FROM [2017$] WHERE [Регион] IS NOT NULL GROUP BY [Регион] ORDER BY COUNT([Регион]) DESC";
                OleDbDataAdapter adapter_region = new OleDbDataAdapter(query_region, con);
                adapter_region.Fill(stat_regions);
                data.Tables.Add(stat_regions);

                OleDbCommand cmd2 = new OleDbCommand("SELECT Sum([Прибыль]) FROM [2017$]", con);
                OleDbDataReader reader = cmd2.ExecuteReader();
                while (reader.Read())
                    stat_income.Content = reader[0].ToString();

                cmd2 = new OleDbCommand("SELECT [Стоимость] FROM [2017$] WHERE [№] IS NOT NULL", con);
                reader = cmd2.ExecuteReader();
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
            region_grid.ItemsSource = data.Tables["Table1"].DefaultView;
        }

        private void check_available_click(object sender, RoutedEventArgs e)
        {
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
                category_for_stock = ch.Content.ToString();

                using (OleDbConnection con = new OleDbConnection(connectionString))
                {
                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    var dataTable = new DataTable();
                    cmd.Parameters.AddWithValue("@arts", arts);
                    string query = string.Format("SELECT * FROM [{0}$] WHERE [Артикул] in (@arts)", ch.Content);
                    OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
                    adapter.Fill(dataTable);
                    data.Tables.Add(dataTable);
                }
                stock_grid.ItemsSource = data.Tables["Table1"].DefaultView;
                break;
            }
            
        }

        private void upload_to_inst(object sender, RoutedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void save_available_click(object sender, RoutedEventArgs e)
        {
            throw new NotImplementedException();
        }
    }
}