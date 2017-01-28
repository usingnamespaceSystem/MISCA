using System;
using System.Windows;
using System.Xml;

namespace MISCA_App
{
    public partial class MainWindow : Window
    {
        void get_curency()
        {
            using (XmlTextReader reader = new XmlTextReader("http://www.cbr.ru/scripts/XML_daily.asp"))
            {
                string USDXml = "";

                while (reader.Read())
                {
                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Element:

                            if (reader.Name == "Valute" && reader.HasAttributes)
                            {
                                while (reader.MoveToNextAttribute())
                                {
                                    if (reader.Name == "ID" && reader.Value == "R01235")
                                    {
                                        reader.MoveToElement();
                                        USDXml = reader.ReadOuterXml();
                                    }
                                }
                            }
                            break;
                    }
                }

                XmlDocument usdXmlDocument = new XmlDocument();
                usdXmlDocument.LoadXml(USDXml);
                XmlNode xmlNode = usdXmlDocument.SelectSingleNode("Valute/Value");
                usd = Convert.ToSingle(xmlNode.InnerText);
                status.Content = "Курс: " + usd.ToString();
            }
        }
    }
}
