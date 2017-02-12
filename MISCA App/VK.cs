using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows;
using VkNet.Model.Attachments;
using VkNet.Model.RequestParams;

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
            var responseImg = Encoding.ASCII.GetString(wc.UploadFile(uploadServer.UploadUrl, AppDomain.CurrentDomain.BaseDirectory + @"\Изображения\main.jpg"));
            var photo = vk.Photo.SaveMarketPhoto(46499802, responseImg);
            wc.Dispose();

            while (count < i)
            {
                    extra_wc = new WebClient();
                    extra_img = Encoding.ASCII.GetString(extra_wc.UploadFile(uploadServer.UploadUrl, AppDomain.CurrentDomain.BaseDirectory + @"\Изображения\" + count + ".jpg"));
                    id = vk.Photo.SaveMarketPhoto(46499802, extra_img);
                    extraPhotos[count - 1] = id.FirstOrDefault().Id.Value;
                    wc.Dispose();
                    count++;
            }

            string descr = string.Empty;

            if (material.Text != string.Empty)
                descr += "Материал: " + material.Text + "\n";
            if (size.Text != string.Empty)
                descr += "Размеры: " + size.Text + "\n";


            var add = vk.Markets.Add(new MarketProductParams
            {
                OwnerId = -46499802,
                CategoryId = 1,
                MainPhotoId = photo.FirstOrDefault().Id.Value,
                Deleted = false,
                Name = (Convert.ToInt32(ws.Cells[rowIdx - 1, 1].Value) + 1).ToString() + " " + name.Text,
                Description = descr,
                Price = Convert.ToDecimal(ws.Cells[rowIdx, 11].Value),
                PhotoIds = extraPhotos

            });

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
