using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net;
using System.Text;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using VkNet.Model.Attachments;
using VkNet.Model.RequestParams;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace MISCA_App
{
    public partial class MainWindow : System.Windows.Window
    {
        long group_id = Convert.ToInt64(Config.Read("group_id", "vk"));

        private void addGoods()
        {
            WebClient extra_wc;
            string extra_img;
            ReadOnlyCollection<Photo> id;

            var uploadServer = _vk.Photo.GetMarketUploadServer(group_id, true, 49, 89, 700);
            var wc = new WebClient();
            ReadOnlyCollection<Photo> photo;
            String responseImg;
            long add;

            try
            {
                responseImg = Encoding.ASCII.GetString(wc.UploadFile(uploadServer.UploadUrl,
                    images_file + @"\main.jpg"));
                photo = _vk.Photo.SaveMarketPhoto(group_id, responseImg);
                wc.Dispose();
            }
            catch (Exception e)
            {
                MessageBox.Show("Произошла ошибка при загрузке главного фото ВК: " + e.Message);
                unexpected_err();
                return;
            }

            if (_isImgAdded)
            {
                _i++;
            }

            try
            {
                while (_count <= _i && _count < 5)
                {
                    extra_wc = new WebClient();
                    extra_img = Encoding.ASCII.GetString(extra_wc.UploadFile(uploadServer.UploadUrl,
                        images_file + _count + ".jpg"));
                    id = _vk.Photo.SaveMarketPhoto(group_id, extra_img);
                    _extraPhotos[_count - 1] = id.FirstOrDefault().Id.Value;
                    wc.Dispose();
                    _count++;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Произошла ошибка при загрузке фото ВК: " + e.Message);
                unexpected_err();
                return;
            }

            string descr = string.Empty;

            if (material.Text != string.Empty)
                descr += "Материал: " + material.Text + "\n";
            if (size.Text != string.Empty)
            {
                if (size.Text.Length > 50)
                    descr += "Размеры: " + "\n" + size.Text + "\n";
                else
                    descr += "Размеры: " + size.Text + "\n";
            }

            if (category.SelectionBoxItem.ToString() == "Лофферы" || category.SelectionBoxItem.ToString() == "Ботинки"
                                                                  || category.SelectionBoxItem.ToString() == "Сумки" ||
                                                                  category.SelectionBoxItem.ToString() == "Кеды"
                                                                  || category.SelectionBoxItem.ToString() ==
                                                                  "Босоножки")
                _catId = 4;

            try
            {
                add = _vk.Markets.Add(new MarketProductParams
                {
                    OwnerId = -group_id,
                    CategoryId = _catId,
                    MainPhotoId = photo.FirstOrDefault().Id.Value,
                    Deleted = false,
                    Name = name.Text + " " + (Convert.ToInt32(_wsheet.Cells[_rowIdx - 1, 1].Value) + 1).ToString(),
                    Description = descr,
                    Price = Convert.ToDecimal(_wsheet.Cells[_rowIdx, 11].Value),
                    PhotoIds = _extraPhotos
                });
            }
            catch (Exception e)
            {
                MessageBox.Show("Произошла ошибка при загрузке товара ВК: " + e.Message);
                unexpected_err();
                return;
            }

            try
            {
                _albumId.Add((Convert.ToInt64((_wsheet.Cells[2, 3] as Range).Value)));
                _vk.Markets.AddToAlbum(ownerId: -group_id, itemId: add, albumIds: _albumId);
            }
            catch (Exception e)
            {
                MessageBox.Show("Произошла ошибка при добавлении товара в подборку: " + e.Message);
                return;
            }

            _nf.Visible = true;
            _nf.Icon = new System.Drawing.Icon(AppDomain.CurrentDomain.BaseDirectory + "bowl.ico");
            _nf.ShowBalloonTip(500, @"¯\_(ツ)_ /¯", "Товар успешно добавлен", System.Windows.Forms.ToolTipIcon.Info);
            uploadServer = null;
            responseImg = null;
            photo = null;
            extra_img = null;
            id = null;
            for (int i1 = 0; i1 < _extraPhotos.Length - 1; i1++)
                _extraPhotos[i1] = 0;
        }

        private void unexpected_err()
        {
            Range rg = (Range) _wsheet.Rows[_rowIdx, Type.Missing];
            rg.Delete(XlDeleteShiftDirection.xlShiftUp);
            _i = 0;
            foreach (FileInfo file in _dirInfo.GetFiles())
            {
                file.Delete();
            }
        }

        private void save_available_Click(object sender, RoutedEventArgs e)
        {
            DataGrid dg = sender as DataGrid;
            string descr = string.Empty;
            bool is_del = false;

            foreach (DataRowView row in dg.Items)
            {
                try
                {
                    if (row["Статус"].ToString() == "0")
                        is_del = true;
                    else { is_del = false; }

                    if (row["Статус"].ToString() == "0" || _arts_edited.Contains(row["Артикул"].ToString()))
                    {

                        var search_items = _vk.Markets.Search(new MarketSearchParams
                        {
                            OwnerId = -group_id,
                            Query = row["Наименование"].ToString()
                        });


                        if (row["Материал"].ToString() != string.Empty)
                            descr += "Материал: " + row["Материал"].ToString() + "\n";

                        if (row["Размеры"].ToString() != string.Empty)
                        {
                            if (row["Размеры"].ToString().Length > 50)
                                descr += "Размеры: " + "\n" + row["Размеры"].ToString() + "\n";
                            else
                                descr += "Размеры: " + row["Размеры"].ToString() + "\n";
                        }

                        foreach (var item in search_items)
                        {
                            var edit = _vk.Markets.Edit(new MarketProductParams
                            {
                                OwnerId = 0-group_id,
                                ItemId = item.Id,
                                Deleted = is_del,
                                Name = row["Наименование"].ToString() + " " + row["Артикул"].ToString(),
                                Description = descr,
                                Price = Convert.ToDecimal(row["Стоимость"]),
                            });
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("Произошла ошибка при снятии с продажи товара" + row["Наименование"].ToString());
                    continue;
                }
            }
            MessageBox.Show("Товары успешно сняты с продажи");
        }
    }
}