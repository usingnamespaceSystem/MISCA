using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Diagnostics;
using System.IO;

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
                cny =8.7;
            };

            clear();

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

            if (isload)
            {
                content = WebControl.ExecuteJavascriptWithResult("document.getElementsByTagName('html')[0].innerHTML");

                Findname();

                Findprice();

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
