using System;
using System.Drawing;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;

namespace MISCA_App
{
    public partial class MainWindow : Window
    {

        private void save_Click(object sender, RoutedEventArgs e)
        {

            wsheet = wbook.Sheets[category.SelectionBoxItem.ToString()];

            while (wsheet.Cells[rowIdx, 1].Value != null)
            {
                rowIdx++;
            }

                if (name.Text.Length == 0 || final_price.Text.Length == 0)
                    MessageBox.Show("Введите имя и цену");

                WebClient wc = new WebClient();

                foreach (CheckingWB cwb in images)
                {
                    if (cwb.checkBox.IsChecked.Value)
                    {
                        if (cwb.main.IsChecked == true)
                        {
                            wc.DownloadFile(cwb.WB.Source, AppDomain.CurrentDomain.BaseDirectory + @"\Изображения\main.jpg");
                            wsheet.Cells[rowIdx, 12].Value = cwb.WB.Source.ToString();
                            ismain = true;
                        }
                        else
                        {
                            i++;
                            wc.DownloadFile(cwb.WB.Source, AppDomain.CurrentDomain.BaseDirectory + @"\Изображения\" + i + ".jpg");
                            wsheet.Cells[rowIdx, img_count+i].Value = cwb.WB.Source.ToString();
                        }
                    }

                }
                if (i==1)
                {
                    MessageBox.Show("Выберите изображения");
                    return;
                }

                if (!ismain)
                {
                    MessageBox.Show("Выберите главное изображение");
                    return;
                }

                wsheet.Cells[rowIdx, 1].Value = Convert.ToInt32(wsheet.Cells[rowIdx - 1, 1].Value) + 1;
                wsheet.Cells[rowIdx, 2].Value = 1;
                wsheet.Cells[rowIdx, 3].Value = link.Text;
                wsheet.Cells[rowIdx, 4].Value = name.Text;
                wsheet.Cells[rowIdx, 5].Value = prod.Text;
                wsheet.Cells[rowIdx, 6].Value = material.Text;
                wsheet.Cells[rowIdx, 7].Value = size.Text;
                wsheet.Cells[rowIdx, 8].Value = price.Text;
                wsheet.Cells[rowIdx, 9].Value = perc.Text;
                wsheet.Cells[rowIdx, 10].Value = ship.Text;
                wsheet.Cells[1, 18].Value = cny;
                wbook.Save();

                 addGoods();

                
                link.Focus();
            }
            //catch (System.Runtime.InteropServices.COMException)
            //{
            //    MessageBox.Show("Произошла ошибка");
            //}

        }
    }

