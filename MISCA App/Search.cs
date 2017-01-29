using System;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows;

namespace MISCA_App
{
    public partial class MainWindow : Window
    {
        /// <summary>
        /// Получение наименования товара из HTML
        /// </summary>
        private void Findname()
        {
            Regex name_tao = new Regex("<span " + "class=\"t-title\" itemprop=\"name\"" + @">(.+)" + "</span>");
            Regex name_world_tmall = new Regex("<h1 data-spm=\"1000983\">\n" + "(.+)" + "\n</h1>");
            Regex script = new Regex("<title>" + @"(.+)" + "-tmall.com");

            MatchCollection matches_name_tao = name_tao.Matches(content);
            MatchCollection matches_name_world_tmall = name_world_tmall.Matches(content);
            MatchCollection matches_script = script.Matches(content);

            if (matches_name_tao.Count > 0)
            {
                name.Text = features_translation(matches_name_tao[0].Groups[1].ToString());
            }
            else if (matches_name_world_tmall.Count > 0)
            {
                name.Text = features_translation(matches_name_world_tmall[0].Groups[1].ToString());
            }
            else if (matches_script.Count > 0)
            {
                name.Text = features_translation(matches_script[0].Groups[1].ToString());
            }
        }

        /// <summary>
        /// Получение цены товара из HTML
        /// </summary>
        private void Findprice()
        {
            Regex price_world_mall = new Regex("<span class=\"tm-price\">" + @"(.+)" + "</span>");
            Regex price_tao = new Regex("<em class=\"tb-rmb\">¥</em>" + @"(.+)" + "</strong>");

            MatchCollection matches_price_world_mall = price_world_mall.Matches(content);
            MatchCollection matches_price_tao = price_tao.Matches(content);

            if (matches_price_world_mall.Count > 0)
            {
                price.Text = matches_price_world_mall[0].Groups[1].ToString();
            }
            else if (matches_price_tao.Count > 0)
            {
                price.Text = matches_price_tao[0].Groups[1].ToString();
            }

            if (price.Text.Contains('.'))
                price.Text = price.Text.Replace('.', ',');
        }

        private void Findseller()
        {
            Regex seller_detail_tmall = new Regex("<li id=\"J_attrBrandName\" title=\"" + @"(.+)" + "\">" + @".+" + "</li>");
            Regex seller_tao = new Regex("<div class=\"tb-shop-name\">" + @"\n<h3><a href=" + @".+" + "title=\"" + @"(.+)" + "\">");

            MatchCollection matches_seller_detall_tmall = seller_detail_tmall.Matches(content);
            MatchCollection matches_seller_tao = seller_tao.Matches(content);

            if (matches_seller_detall_tmall.Count > 0)
            {
                prod.Text = WebUtility.HtmlDecode(Convert.ToString(matches_seller_detall_tmall[0].Groups[1]));
            }
            else if (matches_seller_tao.Count > 0)
            {
                prod.Text = matches_seller_tao[0].Groups[1].ToString();
            }
            else prod.Text = "не определен";
        }

        /// <summary>
        /// Получение материала товара из HTML на TMALL
        /// </summary> 
        private void Findmaterial()
        {
            Regex material1_tmall = new Regex("<li title=\"" + @"(.+)" + "\">Ткань:" + @".+" + "</li>");
            Regex material2_tmall = new Regex("<li title=\" " + @"(.+)" + "\">Материальный компонент: " + @".+" + "</li>");
            Regex material3_tao = new Regex("<li title=\"" + @" (.+)" + "\">\nТкань : " + @".+" + "\n</li>");

            MatchCollection matches_material1_tmall = material1_tmall.Matches(content);
            MatchCollection matches_material2_tmall = material2_tmall.Matches(content);
            MatchCollection matches_material3_tao = material3_tao.Matches(content);

            if (matches_material1_tmall.Count > 0)
            {
                material.Text = features_translation(WebUtility.HtmlDecode(Convert.ToString(matches_material1_tmall[0].Groups[1])).ToString());
            }

            else if (matches_material2_tmall.Count > 0)
            {
                material.Text = features_translation(WebUtility.HtmlDecode(Convert.ToString(matches_material2_tmall[0].Groups[1])).ToString());
            }

            else if (matches_material3_tao.Count > 0)
            {
                material.Text = features_translation(WebUtility.HtmlDecode(Convert.ToString(matches_material3_tao[0].Groups[1])).ToString());
            }
        }


        /// <summary>
        /// Получение изображений товара из HTML
        /// </summary>
        private void Get_images()
        {
            Regex descURLtao = new Regex("descUrl: \"" + @"(.{140,150})" + "\"");
            Regex descURLtmall = new Regex("{\"descUrl\":\"" + @"(.{155,165})" + "\",");
            MatchCollection matchURLtao = descURLtao.Matches(content);
            MatchCollection matchURLtmall = descURLtmall.Matches(content);
            WebClient wb = new WebClient();

            if(matchURLtao.Count >0)
                imagesHidden = wb.DownloadString(new Uri("http:"+matchURLtao[0].Groups[1].ToString()));
            if (matchURLtmall.Count > 0)
                imagesHidden = wb.DownloadString(new Uri("http:" + matchURLtmall[0].Groups[1].ToString()));
           
            Regex img_tmall = new Regex("data-ks-lazyload=\"https:" + "//img.alicdn.com/" + @"(.{50,80})" + ".jpg\">");
            Regex img_tmall2 = new Regex("//img.alicdn.com/" + @"(.{50,80})" + ".jpg\"");
            Regex img_tao = new Regex("//img.alicdn.com/" + @"(.{50,80})" + ".gif\"");

            MatchCollection matches_img_tmall = img_tmall.Matches(imagesHidden);
            MatchCollection matches_img_tao = img_tao.Matches(imagesHidden);
            MatchCollection matches_img_tmall2 = img_tmall2.Matches(imagesHidden);

            if (matches_img_tmall.Count > 0)
            {
                for (int n = 0; n != matches_img_tmall.Count; n++)
                {
                    CheckingWB img_parser = new CheckingWB();
                    images.Add(img_parser);
                    img_parser.Name = "img" + (n + 1).ToString();

                    img_parser.WB.Source = new Uri("https://img.alicdn.com/" + matches_img_tmall[n].Groups[1].ToString() + ".jpg");

                    img.Children.Add(img_parser);
                }
            }
            if (matches_img_tmall2.Count > 0)
            {
                for (int n = 0; n != matches_img_tmall2.Count; n++)
                {
                    CheckingWB img_parser = new CheckingWB();
                    images.Add(img_parser);
                    img_parser.Name = "img" + (n + 1).ToString();

                    img_parser.WB.Source = new Uri("https://img.alicdn.com/" + matches_img_tmall2[n].Groups[1].ToString() + ".jpg");

                    img.Children.Add(img_parser);
                }
            }

            if (matches_img_tao.Count > 0)
            {
                for (int n = 0; n != matches_img_tao.Count; n++)
                {
                    CheckingWB img_parser = new CheckingWB();
                    images.Add(img_parser);
                    img_parser.Name = "img" + (n + 1).ToString();

                    img_parser.WB.Source = new Uri("https://img.alicdn.com/" + matches_img_tao[n].Groups[1].ToString() + ".gif");

                    img.Children.Add(img_parser);

                }
            }


        }
    }
}
