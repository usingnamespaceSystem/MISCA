using Awesomium.Windows.Controls;
using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

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
            };

            clear();

            foreach (FileInfo file in dirInfo.GetFiles())
            {
                file.Delete();
            }
            WebControl_promo.Visibility = Visibility.Visible;
            TranslateText();
            
        }

        private void link_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
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

                    WebControl_promo.Visibility = Visibility.Hidden;
                };
                
                content = WebControl.ExecuteJavascriptWithResult("document.getElementsByTagName('html')[0].innerHTML");

                Findname();

                Findseller();

                Findmaterial();

                Get_images();

                isload = false;

                perc.Text = "7";
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {

            foreach (FileInfo file in dirInfo.GetFiles())
            {
                file.Delete();
            }

            wb.Close(true);
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

        private void forward_Click(object sender, RoutedEventArgs e)
        {
            WebControl.GoForward();
        }

        private void back_Click(object sender, RoutedEventArgs e)
        {
            WebControl.GoBack();
        }

        private void price_changed(object sender, TextChangedEventArgs e)
        {
            if (price.Text != String.Empty && perc.Text != String.Empty && ship.Text != String.Empty)
            {
                final_price.Text = Math.Round(((Convert.ToDouble(price.Text) * (1.07 + Convert.ToDouble(perc.Text) / 100.0) + 20) * cny + Convert.ToDouble(ship.Text)), 0).ToString();
                price1.Content = "(" + Math.Round(((Convert.ToDouble(price.Text) * 1.07 + 20) * cny + Convert.ToDouble(ship.Text)), 0).ToString() + ")";
            }

        }

        private void ship_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (price.Text != String.Empty && perc.Text != String.Empty && ship.Text != String.Empty)
            {
                final_price.Text = Math.Round(((Convert.ToDouble(price.Text) * (1.07 + Convert.ToDouble(perc.Text) / 100.0) + 20) * cny + Convert.ToDouble(ship.Text)), 0).ToString();
                price1.Content = "(" + Math.Round(((Convert.ToDouble(price.Text) * 1.07 + 20) * cny + Convert.ToDouble(ship.Text)), 0).ToString() + ")";
            }

        }

        private void perc_changed(object sender, TextChangedEventArgs e)
        {
            if (price.Text != String.Empty && perc.Text != String.Empty && ship.Text != String.Empty)
            {
                final_price.Text = Math.Round(((Convert.ToDouble(price.Text) * (1.07 + Convert.ToDouble(perc.Text) / 100.0) + 20) * cny + Convert.ToDouble(ship.Text)), 0).ToString();
                price1.Content = "(" + Math.Round(((Convert.ToDouble(price.Text) * 1.07 + 20) * cny + Convert.ToDouble(ship.Text)), 0).ToString() + ")";
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

        private void perc_KeyDown(object sender, KeyEventArgs e)
        {
            if (!((e.Key.GetHashCode() >= 34) && (e.Key.GetHashCode() <= 43)) && !((e.Key.GetHashCode() >= 74) && (e.Key.GetHashCode() <= 83)) && e.Key.GetHashCode() == 73)
            {
                e.Handled = true;
            }
        }

        private void ship_KeyDown(object sender, KeyEventArgs e)
        {
            if (!((e.Key.GetHashCode() >= 34) && (e.Key.GetHashCode() <= 43)) && !((e.Key.GetHashCode() >= 74) && (e.Key.GetHashCode() <= 83)) && e.Key.GetHashCode() == 73)
            {
                e.Handled = true;
            }
        }

        private void final_price_KeyDown(object sender, KeyEventArgs e)
        {
            if (!((e.Key.GetHashCode() >= 34) && (e.Key.GetHashCode() <= 43)) && !((e.Key.GetHashCode() >= 74) && (e.Key.GetHashCode() <= 83)) && e.Key.GetHashCode() == 73)
            {
                e.Handled = true;
            }
        }

    }
}
