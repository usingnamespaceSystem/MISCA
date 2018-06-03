using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Diagnostics;
using System.IO;
using System.Windows.Data;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Reflection;

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
                cny = 8.7;
                MessageBox.Show("Курс не получен");
            };

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
            string promotion = string.Empty;

            if (isload)
            {
                WebControl_promo.Source = new Uri(link.Text);
                WebControl_promo.LoadingFrameComplete += (obj, evt) =>
                {
                    string script = @"(function() { for (var i in g_config.promotion.promoData) { return g_config.promotion.promoData[i][0].price } }())";
                    promotion = WebControl_promo.ExecuteJavascriptWithResult(script);

                    if (promotion != "undefined")
                        price.Text = promotion.Trim('"').Replace('.', ',');

                    if (promotion == "undefined")
                        Findprice();

                    if (price.Text.Contains("-"))
                        price.Text = price.Text.Split('-')[1].Trim();

                    if (price.Text != "")
                        perc.Text = Convert.ToString(Math.Round(2500 / Convert.ToDecimal(price.Text.Replace('.', ',')), 0));

                    if (perc.Text == "0" || perc.Text == "")
                    {
                        perc.Text = "20";
                    }

                    WebControl_promo.Visibility = Visibility.Hidden;
                };

                content = WebControl.ExecuteJavascriptWithResult("document.getElementsByTagName('html')[0].innerHTML");

                Findname();

                Findseller();

                Findmaterial();

                Get_images();

                isload = false;

                auto_category();
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            foreach (FileInfo file in dirInfo.GetFiles())
            {
                file.Delete();
            }

            try
            {
                wbook.Close(true);
            }
            catch (Exception)
            { MessageBox.Show("Не удалось закрыть Excel"); }


            app.Quit();

            try
            {
                foreach (var process in Process.GetProcessesByName("EXCEL"))
                {
                    process.Kill();
                }
            }
            catch { }
        }

        private void img_added_Click(object sender, RoutedEventArgs e)
        {
            if (i > 0 && i < 5)
            {
                isImgAdded = true;
                img_checking_count.Content = Convert.ToInt32(img_checking_count.Content) + 1;
            }
        }


        private void img_deleted_Click(object sender, RoutedEventArgs e)
        {
            if (i > 0 && i < 5)
            {
                isImgAdded = false;
                img_checking_count.Content = Convert.ToInt32(img_checking_count.Content) - 1;
            }
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
            for (int i = 1; i <= parameters.Count; i++)
            {
                size_table.Columns[i - 1].Header = parameters[i - 1];
            }
            //size_table.Items.Add(new ListCollectionView(parameters));
            SizeRow new_row = new SizeRow() { field1 = string.Empty, field2 = string.Empty, field3 = string.Empty, field4 = string.Empty, field5 = string.Empty };
            SizeRowCollection.Add(new_row);

            isSizeInTable = true;
        }


        private void back_Click(object sender, RoutedEventArgs e)
        {
            WebControl.GoBack();
        }

        private void change_price(object sender, TextChangedEventArgs e)
        {
            if (price.Text != String.Empty && perc.Text != String.Empty)
            {
                try
                {
                    final_price.Text = Math.Round(((Convert.ToDouble(price.Text) * (1.07 + Convert.ToDouble(perc.Text) / 100.0) + 20) * cny + Convert.ToDouble(ship.Content)), 0).ToString();
                    income.Content = (Convert.ToDouble(final_price.Text) - Math.Round(((Convert.ToDouble(price.Text) * 1.07 + 20) * cny + Convert.ToDouble(ship.Content)), 0)).ToString();
                }
                catch (Exception ex)
                { MessageBox.Show("Произошла ошибка при изменении цены: " + ex.Message); }
            }
        }

        private void change_weight(object sender, TextChangedEventArgs e)
        {
            if (weight.Text != String.Empty)
            {
                try
                {
                    int agents_comission = 0;
                    //комиссия посредника считается либо за вес, либо единожды за посылку вцелом
                    if (agent_row.Columns[3].Value)
                    {
                        agents_comission = weight.Text * agent_row.Columns[3].Value;
                    }
                    else
                    {
                        agents_comission = weight.Text * agent_row.Columns[4].Value;
                    }
                    //комиссия посрдника считается либо за вес, либо единожды за посылку вцелом
                    ship.Content = weight.Text * agent_row.Columns[5].Value + agents_comission;
                }
                catch (Exception ex)
                { MessageBox.Show("Произошла ошибка при изменении веса: " + ex.Message); }
            }
        }


        private void category_DropDownClosed(object sender, System.EventArgs e)
        {
            foreach (Microsoft.Office.Interop.Excel.Worksheet sh in wbook.Worksheets)
            {
                if (sh.Name != category.SelectionBoxItem.ToString())
                { continue; }

                weight.Text = Convert.ToInt32((sh.Cells[2, 5] as Microsoft.Office.Interop.Excel.Range).Value);
                int agents_comission = 0;
                //комиссия посредника считается либо за вес, либо единожды за посылку вцелом
                if (agent_row.Columns[3].Value)
                {
                    agents_comission = weight.Text * agent_row.Columns[3].Value;
                }
                else
                {
                    agents_comission = weight.Text * agent_row.Columns[4].Value;
                }
                //комиссия посредника считается либо за вес, либо единожды за посылку вцелом
                ship.Content = weight.Text * agent_row.Columns[5].Value + agents_comission;
                return;
            }
        }

        private void agent_DropDownClosed(object sender, System.EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Worksheet wsheet_agents = wbook_agents.Worksheets[1];
            foreach (Microsoft.Office.Interop.Excel.Range row in wsheet_agents.UsedRange.Rows)
            {
                if (row.Columns[6].Text == agent.SelectionBoxItem.ToString())
                {
                    agent_row = row;
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
            if (!((e.Key.GetHashCode() >= 34) && (e.Key.GetHashCode() <= 43)) && !((e.Key.GetHashCode() >= 74) && (e.Key.GetHashCode() <= 83)) && e.Key.GetHashCode() == 73)
            {
                e.Handled = true;
            }
        }

        private void select_all_category(object sender, RoutedEventArgs e)
        {
            foreach (CheckBox ch in category_panel.Children)
            {
                ch.IsChecked = true;
            }
        }

        private void unselect_all_category(object sender, RoutedEventArgs e)
        {
            foreach (CheckBox ch in category_panel.Children)
            {
                ch.IsChecked = false;
            }
        }

        private void check_available_click(object sender, RoutedEventArgs e)
        {
            string reply = string.Empty;
            //int link_column = 3;
            stock_grid.ItemsSource = StockRowCollection;

            for (int i = 1; i <= last_column; i++)
            {
                string header = wbook.Worksheets[1].Cells[1, i].Text;
                stock_grid.Columns[i - 1].Header = header;
                stock_grid.Columns[i - 1].MaxWidth = 200;
                if (header == "Ссылка")
                {
                    link_column = i;
                }
            }

            foreach (CheckBox ch in category_panel.Children)
            {
                Microsoft.Office.Interop.Excel.Worksheet wsheet = wbook.Worksheets[ch.Content];
                if (ch.IsChecked == true)
                {
                    foreach (Microsoft.Office.Interop.Excel.Range row in wsheet.UsedRange.Rows)
                    {
                        if (row.Row == 1 )
                            continue;

                        using (WebClient client = new WebClient())
                        {
                            try
                            {
                                client.Encoding = System.Text.Encoding.GetEncoding("GB2312");
                                reply = client.DownloadString(row.Columns[link_column].Text);
                            }
                            catch (WebException ex)
                            {
                                if (ex.Status == WebExceptionStatus.ProtocolError && ex.Response != null)
                                {
                                    StockRowCollection.Add(new StockRow()
                                    {
                                        stock_article = row.Columns[1].Text,
                                        stock_status = row.Columns[2].Text,
                                        stock_name = row.Columns[3].Text,
                                        stock_link = row.Columns[4].Text,
                                        stock_seller = row.Columns[5].Text,
                                        stock_material = row.Columns[6].Text,
                                        stock_size = row.Columns[7].Text,
                                        stock_price = row.Columns[8].Text,
                                        stock_percent = row.Columns[9].Text,
                                        stock_shipping = row.Columns[10].Text,
                                        stock_summary = row.Columns[11].Text
                                    });
                                    continue;
                                }
                            }
                        }

                        if (reply.Contains("此宝贝已下架") || reply.Contains("此商品已下架") || reply.Contains("您查看的宝贝不存在"))
                        {
                            StockRowCollection.Add(new StockRow()
                            {
                                stock_article = row.Columns[1].Text,
                                stock_status = row.Columns[2].Text,
                                stock_name = row.Columns[3].Text,
                                stock_link = row.Columns[4].Text,
                                stock_seller = row.Columns[5].Text,
                                stock_material = row.Columns[6].Text,
                                stock_size = row.Columns[7].Text,
                                stock_price = row.Columns[8].Text,
                                stock_percent = row.Columns[9].Text,
                                stock_shipping = row.Columns[10].Text,
                                stock_summary = row.Columns[11].Text
                            });
                            reply = string.Empty;
                        }
                    }
                }
            }
        }
    }
}
