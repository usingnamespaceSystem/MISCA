using System;
using System.Drawing;
using System.Net;
using System.Windows;
using System.Windows.Forms;

namespace MISCA_App
{
    public partial class MainWindow : Window
    {
        private void save_Click(object sender, RoutedEventArgs e)
        {
            i = 1;

            WebClient wc = new WebClient();

            foreach (CheckingWB cwb in images)
            {
                if (cwb.checkBox.IsChecked.Value)
                {
                    if (cwb.main.IsChecked == true)
                        wc.DownloadFile(cwb.WB.Source, AppDomain.CurrentDomain.BaseDirectory + @"\Изображения\main.jpg");
                    else
                    {
                        wc.DownloadFile(cwb.WB.Source, AppDomain.CurrentDomain.BaseDirectory + @"\Изображения\" + i + ".jpg");
                        i++;
                    }
                }

            }

            ws = wb.Sheets[category.SelectionBoxItem.ToString()];

            while (ws.Cells[rowIdx, 1].Value != null)
            {
                rowIdx++;
            }
           
            ws.Cells[rowIdx, 1].Value = Convert.ToInt32(ws.Cells[rowIdx - 1, 1].Value) + 1;
            ws.Cells[rowIdx, 2].Value = 1;
            ws.Cells[rowIdx, 3].Value = WebControl.Source.ToString();
            ws.Cells[rowIdx, 4].Value = name.Text;
            ws.Cells[rowIdx, 5].Value = prod.Text;
            ws.Cells[rowIdx, 6].Value = material.Text;
            ws.Cells[rowIdx, 7].Value = size.Text;
            ws.Cells[rowIdx, 8].Value = price.Text;
            ws.Cells[rowIdx, 9].Value = perc.Text;
            ws.Cells[rowIdx, 10].Value = ship.Text;
            ws.Cells[1, 14].Value = cny;
            wb.Save();

            addGoods();

            nf.Visible = true;
            nf.Icon = new Icon(AppDomain.CurrentDomain.BaseDirectory + "bowl.ico");
            nf.ShowBalloonTip(500, @"¯\_(ツ)_ /¯", "Товар успешно добавлен", ToolTipIcon.Info);

            link.Focus();

        }
    }
}
