using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows;
using VkNet.Model.Attachments;
using VkNet.Model.RequestParams;
using System.Drawing;

namespace MISCA_App
{
    public partial class MainWindow : Window
    {
        private void addGoods()
        {
            WebClient extra_wc;
            string extra_img;
            ReadOnlyCollection<Photo> id;

            var uploadServer = vk.Photo.GetMarketUploadServer(46499802, true, 49, 89, 700);
            var wc = new WebClient();
            System.Collections.ObjectModel.ReadOnlyCollection<Photo> photo;
            String responseImg;

            try
            {

                responseImg = Encoding.ASCII.GetString(wc.UploadFile(uploadServer.UploadUrl, AppDomain.CurrentDomain.BaseDirectory + @"\Изображения\main.jpg"));
                photo = vk.Photo.SaveMarketPhoto(46499802, responseImg);
                wc.Dispose();
            }
            catch (Exception e)
            {
                MessageBox.Show("Произошла ошибка при загрузке главного фото ВК: " + e.Message);
                Microsoft.Office.Interop.Excel.Range rg = (Microsoft.Office.Interop.Excel.Range)wsheet.Rows[rowIdx, Type.Missing];
                rg.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
                return;
            }

                if (isImgAdded) { i++; }

            try
            {
                while (count <= i && count < 5)
                {
                    extra_wc = new WebClient();
                    extra_img = Encoding.ASCII.GetString(extra_wc.UploadFile(uploadServer.UploadUrl, AppDomain.CurrentDomain.BaseDirectory + @"\Изображения\" + count + ".jpg"));
                    id = vk.Photo.SaveMarketPhoto(46499802, extra_img);
                    extraPhotos[count - 1] = id.FirstOrDefault().Id.Value;
                    wc.Dispose();
                    count++;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Произошла ошибка при загрузке фото ВК: " + e.Message);
                Microsoft.Office.Interop.Excel.Range rg = (Microsoft.Office.Interop.Excel.Range)wsheet.Rows[rowIdx, Type.Missing];
                rg.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
                return;
            }

            string descr = string.Empty;

            if (material.Text != string.Empty)
                descr += "Материал: " + material.Text + "\n";
            if (size.Text != string.Empty)
                descr += "Размеры: " + size.Text + "\n";

            try
            {
                var add = vk.Markets.Add(new MarketProductParams
                {
                    OwnerId = -46499802,
                    CategoryId = 1,
                    MainPhotoId = photo.FirstOrDefault().Id.Value,
                    Deleted = false,
                    Name = name.Text + " " + (Convert.ToInt32(wsheet.Cells[rowIdx - 1, 1].Value) + 1).ToString(),
                    Description = descr,
                    Price = Convert.ToDecimal(wsheet.Cells[rowIdx, 11].Value),
                    PhotoIds = extraPhotos

                });
            }
            catch (Exception e)
            {
                MessageBox.Show("Произошла ошибка при загрузке товара ВК: " + e.Message);
                Microsoft.Office.Interop.Excel.Range rg = (Microsoft.Office.Interop.Excel.Range)wsheet.Rows[rowIdx, Type.Missing];
                rg.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);
                return;
            }
            nf.Visible = true;
            nf.Icon = new Icon(AppDomain.CurrentDomain.BaseDirectory + "bowl.ico");
            nf.ShowBalloonTip(500, @"¯\_(ツ)_ /¯", "Товар успешно добавлен", System.Windows.Forms.ToolTipIcon.Info);

            uploadServer = null;
            responseImg = null;
            photo = null;
            extra_img = null;
            id = null;
            for (int i1 = 0; i1 < extraPhotos.Length - 1; i1++)
                extraPhotos[i1] = 0;

        }
    }
}
