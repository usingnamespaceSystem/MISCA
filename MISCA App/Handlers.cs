using System;
using System.Threading.Tasks;
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
            clear();
            WebControl.Dispatcher.BeginInvoke(new Translation(TranslateText));
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

        private async void WebControl_LoadingFrameComplete(object sender, Awesomium.Core.FrameEventArgs e)
        {
            if (load)
            {
                await Task.Factory.StartNew<int>(
                                        () =>
                                        {
                                            WebControl.Dispatcher.BeginInvoke(new Action(delegate ()
                                            {
                                                WebControl.ExecuteJavascriptWithResult("window.scrollTo(0,1500)");                         

                                                WebControl.ExecuteJavascriptWithResult("window.scrollTo(0,4000)");

                                                Task.Delay(500).Wait();

                                                content = WebControl.ExecuteJavascriptWithResult("document.getElementsByTagName('html')[0].innerHTML");

                                                Get_images();
                                            }));
                                            return 1;
                                        },
                                        TaskCreationOptions.LongRunning);

                status.Content = "Загрузка завершена";

                Findname();

                Findprice();

                Findseller();

                Findmaterial();

                load = false;
            }
        }



        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            wb.Close(true);
            app.Quit();
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
                final_price.Text = Math.Round(((Convert.ToDouble(price.Text) * 0.15 * (1.0 + Convert.ToDouble(perc.Text) / 100.0) + 10) * usd + Convert.ToDouble(ship.Text)), 0).ToString();
        }

        private void ship_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (price.Text != String.Empty && perc.Text != String.Empty && ship.Text != String.Empty)
                final_price.Text = Math.Round(((Convert.ToDouble(price.Text) * 0.15 * (1 + Convert.ToDouble(perc.Text) / 100) + 10) * usd + Convert.ToDouble(ship.Text)), 0).ToString();
        }

        private void WebControl_LoadingFrame(object sender, Awesomium.Core.LoadingFrameEventArgs e)
        {
            status.Content = "Загрузка страницы...";
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Auth();
        }

        private void perc_changed(object sender, TextChangedEventArgs e)
        {
            if (price.Text != String.Empty && perc.Text != String.Empty && ship.Text != String.Empty)
                final_price.Text = Math.Round(((Convert.ToDouble(price.Text) * 0.15 * (1 + Convert.ToDouble(perc.Text) / 100) + 10) * usd + Convert.ToDouble(ship.Text)), 0).ToString();
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
