using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace MISCA_App
{
    public partial class MainWindow : Window
    {
        void clear()
        {
            nf.Visible = false;

            img_checking_count.Content = "0";

            isload = true;

            imagesHidden = string.Empty;

            images.RemoveRange(0, (images.Count != 0) ? images.Count - 1 : 0);

            images.Clear();

            for(int i=img.Children.Count-1; i>=0;i--)
            {
                img.Children.RemoveAt(i);
            }

            foreach (TextBox c in FindVisualChildren<TextBox>(this))
                if (!c.Name.Contains("link") && !c.Name.Contains("ship"))
                    c.Text = "";

            foreach (FileInfo file in dirInfo.GetFiles())
            {
                file.Delete();
            }

            nf.Visible = false;
        }

        public static IEnumerable<T> FindVisualChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj != null)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
                {
                    DependencyObject child = VisualTreeHelper.GetChild(depObj, i);
                    if (child != null && child is T)
                    {
                        yield return (T)child;
                    }

                    foreach (T childOfChild in FindVisualChildren<T>(child))
                    {
                        yield return childOfChild;
                    }
                }
            }
        }

    }
}
