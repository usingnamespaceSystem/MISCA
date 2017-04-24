using System;
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
            i = 0;
            count = 1;
            rowIdx = 1;
            img_count = 12;

            nf.Visible = false;
            isImgAdded = false;
            isload = true;

            img_checking_count.Content = "0";

            imagesHidden = string.Empty;

            images.RemoveRange(0, (images.Count != 0) ? images.Count - 1 : 0);

            images.Clear();
            album_id.Clear();

            for (int i=img.Children.Count-1; i>=0;i--)
                img.Children.RemoveAt(i);

            foreach (TextBox c in FindVisualChildren<TextBox>(this))
                if (!c.Name.Contains("link") && !c.Name.Contains("ship"))
                    c.Text = "";

            foreach (FileInfo file in dirInfo.GetFiles())
                file.Delete();
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
