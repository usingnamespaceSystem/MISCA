using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows;
using VkNet.Model.Attachments;
using VkNet.Model.RequestParams;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace MISCA_App
{
    public partial class MainWindow : System.Windows.Window
    {
        private void addGoods()
        {
            WebClient extra_wc;
            string extra_img;
            ReadOnlyCollection<Photo> id;

            var uploadServer = _vk.Photo.GetMarketUploadServer(46499802, true, 49, 89, 700);
            var wc = new WebClient();
            ReadOnlyCollection<Photo> photo;
            String responseImg;
            long add;

            try
            {
                responseImg = Encoding.ASCII.GetString(wc.UploadFile(uploadServer.UploadUrl,
                    AppDomain.CurrentDomain.BaseDirectory + @"\Изображения\main.jpg"));
                photo = _vk.Photo.SaveMarketPhoto(46499802, responseImg);
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
                        AppDomain.CurrentDomain.BaseDirectory + @"\Изображения\" + _count + ".jpg"));
                    id = _vk.Photo.SaveMarketPhoto(46499802, extra_img);
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
                    OwnerId = -46499802,
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
                _vk.Markets.AddToAlbum(ownerId: -46499802, itemId: add, albumIds: _albumId);
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
    }
}