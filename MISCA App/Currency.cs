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
                string CNYXml = "";

                while (reader.Read())
                {
                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Element:

                            if (reader.Name == "Valute" && reader.HasAttributes)
                            {
                                while (reader.MoveToNextAttribute())
                                {
                                    if (reader.Name == "ID" && reader.Value == "R01375")
                                    {
                                        reader.MoveToElement();
                                        CNYXml = reader.ReadOuterXml();
                                    }
                                }
                            }
                            break;
                    }
                }

                XmlDocument cnyXmlDocument = new XmlDocument();
                cnyXmlDocument.LoadXml(CNYXml);
                XmlNode xmlNode = cnyXmlDocument.SelectSingleNode("Valute/Value");
                cny = Convert.ToDouble(xmlNode.InnerText) / 10.0;
                status.Content = "Курс: " + cny.ToString();
            }
        }
    }
}
