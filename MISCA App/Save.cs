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
            _wsheet = _wbook.Sheets[category.SelectionBoxItem.ToString()];

            while (_wsheet.Cells[_rowIdx, 1].Value != null)
                _rowIdx++;

            if (name.Text.Length == 0 || final_price.Text.Length == 0)
                MessageBox.Show("Введите имя и цену");

            WebClient wc = new WebClient();

            foreach (CheckingWB cwb in _images)
            {
                if (cwb.checkBox.IsChecked.Value)
                {
                    if (cwb.main.IsChecked == true)
                    {
                        wc.DownloadFile(cwb.WB.Source,
                            $@"{AppDomain.CurrentDomain.BaseDirectory}\Изображения\main.jpg");
                        _wsheet.Cells[_rowIdx, 12].Value = cwb.WB.Source.ToString();
                        _ismain = true;
                    }
                    else
                    {
                        _i++;
                        wc.DownloadFile(cwb.WB.Source,
                            $@"{AppDomain.CurrentDomain.BaseDirectory}\Изображения\{_i}.jpg");
                        _wsheet.Cells[_rowIdx, _imgCount + _i].Value = cwb.WB.Source.ToString();
                    }
                }
            }

            if (_i == 1)
            {
                MessageBox.Show("Выберите изображения");
                return;
            }

            if (!_ismain)
            {
                MessageBox.Show("Выберите главное изображение");
                return;
            }

            _wsheet.Cells[_rowIdx, 1].Value = Convert.ToInt32(_wsheet.Cells[_rowIdx - 1, 1].Value) + 1;
            _wsheet.Cells[_rowIdx, 2].Value = 1;
            _wsheet.Cells[_rowIdx, 3].Value = link.Text;
            _wsheet.Cells[_rowIdx, 4].Value = name.Text;
            _wsheet.Cells[_rowIdx, 5].Value = prod.Text;
            _wsheet.Cells[_rowIdx, 6].Value = material.Text;

            if (_isSizeInTable == true)
            {
                string size_str = string.Empty;
                for (int i = 1; i <= size_table.Columns.Count; i++)
                {
                    for (int j = 1; j <= size_table.Items.Count; j++)
                    {
                        //ПОлучение значения из ячйки
                        size_str +=
                            $"{size_table.Columns[i - 1].Header}-{(size_table.Columns[i - 1].GetCellContent(size_table.Items[j - 1]) as TextBlock)?.Text}см, ";
                    }

                    size_str = $"{size_str.Remove(size_str.Length - 1, 1)}\n";
                }

                _wsheet.Cells[_rowIdx, 7].Value = size_str;
            }
            else
                _wsheet.Cells[_rowIdx, 7].Value = size.Text;

            _wsheet.Cells[_rowIdx, 8].Value = price.Text;
            _wsheet.Cells[_rowIdx, 9].Value = perc.Text;
            _wsheet.Cells[_rowIdx, 10].Value = ship.Content;
            _wsheet.Cells[1, 18].Value = _cny;
            _wbook.Save();
            addGoods();
            link.Focus();
        }

        //catch (System.Runtime.InteropServices.COMException)
        //{
        //    MessageBox.Show("Произошла ошибка");
        //}
    }
}