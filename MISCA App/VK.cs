using System;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows;
using VkNet.Model.RequestParams;

namespace MISCA_App
{
    public partial class MainWindow : Window
    {
        private void addGoods()
        {
            var uploadServer = vk.Photo.GetMarketUploadServer(46499802, true, 0, 0, 600);
            var wc = new WebClient();
            var responseImg = Encoding.ASCII.GetString(wc.UploadFile(uploadServer.UploadUrl, AppDomain.CurrentDomain.BaseDirectory + @"\Изображения\main.jpg"));
            var photo = vk.Photo.SaveMarketPhoto(46499802, responseImg);
            wc.Dispose();

            while(count < i)
            { 
                var extra_wc = new WebClient();
                var img = Encoding.ASCII.GetString(extra_wc.UploadFile(uploadServer.UploadUrl, AppDomain.CurrentDomain.BaseDirectory + @"\Изображения\"+ count + ".jpg"));
                var id = vk.Photo.SaveMarketPhoto(46499802, img);
                extraPhotos[count - 1] = id.FirstOrDefault().Id.Value;
                wc.Dispose();
                count++;
            }

            string descr = string.Empty;
            if (size.Text != string.Empty)
                descr += "Размеры: " + size.Text + "\n";
            if (material.Text != string.Empty)
                descr += "Материал: " + material.Text + "\n";

            var add = vk.Markets.Add(new MarketProductParams
            {
                OwnerId = -46499802,
                CategoryId = 1,
                MainPhotoId = photo.FirstOrDefault().Id.Value,
                Deleted = false,
                Name = (Convert.ToInt32(ws.Cells[rowIdx - 1, 1].Value) + 1).ToString() + " " + name.Text,
                Description = descr,
                Price = Convert.ToDecimal(final_price.Text),
                PhotoIds = extraPhotos

            });
        }
    }
}
