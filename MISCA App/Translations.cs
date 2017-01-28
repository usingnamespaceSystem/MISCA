using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;

namespace MISCA_App
{
    public partial class MainWindow : Window
    {

        /// <summary>
        /// Перевод входной страницы
        /// </summary>
        public void TranslateText()
        {
            Regex item_tao = new Regex(@"(https://item.taobao.com/item)" + @".+" + @"\&" + @"(id=)" + @"(\d+)");
            Regex world_tao = new Regex(@"(https://world.taobao.com/item/)" + @"(\d+)" + "(.htm)" + ".+");
            Regex detail_tmall = new Regex(@"(https://detail.tmall.com/item\.htm\?)" + "(id=)" + @"(\d+)");
            Regex detail_tmall_long = new Regex(@"(https://detail.tmall.com/item)" + @".+" + @"\&" + @"(id=)" + @"(\d+)" + @".+");
            Regex world_tmall = new Regex(@"(https://world.tmall.com/item/)" + @"(\d+)" + @".+");

            MatchCollection match_item_tao = item_tao.Matches(link.Text);
            MatchCollection match_world_tao = world_tao.Matches(link.Text);
            MatchCollection match_detail_tmall = detail_tmall.Matches(link.Text);
            MatchCollection match_detail_tmall_long = detail_tmall_long.Matches(link.Text);
            MatchCollection match_world_tmall_long = world_tmall.Matches(link.Text);

            String url_ready = String.Empty;

            if (match_item_tao.Count > 0)
                url_ready = match_item_tao[0].Groups[1] + ".htm?" + match_item_tao[0].Groups[2] + match_item_tao[0].Groups[3];
            else if (match_world_tao.Count > 0)
                url_ready = match_world_tao[0].Groups[1] + "" + match_world_tao[0].Groups[2] + match_world_tao[0].Groups[3] + "/";
            else if (match_detail_tmall_long.Count > 0)
                url_ready = match_detail_tmall_long[0].Groups[1] + ".htm?" + match_detail_tmall_long[0].Groups[2] + match_detail_tmall_long[0].Groups[3];
            else if (match_detail_tmall.Count > 0)
                url_ready = link.Text;
            else if (match_world_tmall_long.Count > 0)
                url_ready = link.Text;
            else url_ready = link.Text;

            string url = String.Format("https://z5h64q92x9.net/tr-start?ui=ru&url={0}&lang=zh-ru", url_ready);

            WebControl.Source = new Uri(url);
        }
        /// <summary>
        /// Перевод характеристик товара
        /// </summary>
        /// <param name="feature"></param>
        /// <returns></returns>
        string features_translation(string feature)
        {
            string result = translate(feature, "zh", "ru");
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(result);
            HtmlAgilityPack.HtmlNodeCollection translateNode = doc.DocumentNode.SelectNodes("//span[@id='result_box']");

            if (translateNode.Count != 0)
            {
                return feature = translateNode[0].InnerText;
            }
            else
            {
                return feature += "  ОШИБКА";
            }
        }

        /// <summary>
        /// Перевод характеристик товара
        /// </summary>
        /// <param name="word">Характеристика</param>
        /// <param name="SL">Язык источника</param>
        /// <param name="DL">Язык перевода</param>
        /// <returns></returns>
        public string translate(string word, string SL, string DL)
        {
            CookieContainer cookies = new CookieContainer();
            string result;
            ServicePointManager.Expect100Continue = false;
            var request = (HttpWebRequest)WebRequest.Create("https://translate.google.ru/?sl=" + SL + "&tl=" + DL + "&q=" + word);

            request.CookieContainer = cookies;
            request.Credentials = CredentialCache.DefaultCredentials;
            request.Proxy.Credentials = CredentialCache.DefaultCredentials;
            request.Method = "PUT";
            request.ContentType = "application/x-www-form-urlencoded";
            request.UserAgent = @"Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.8.0.4) Gecko/20060508 Firefox/1.5.0.4";

            using (var requestStream = request.GetRequestStream())
            using (var responseStream = request.GetResponse().GetResponseStream())
            using (var reader = new StreamReader(responseStream, Encoding.UTF8))
            {
                result = reader.ReadToEnd();
            }

            return result;
        }
    }
}
