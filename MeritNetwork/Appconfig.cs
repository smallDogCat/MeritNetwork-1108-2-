using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace MeritNetwork
{
    class Appconfig
    {
        public static void UpdateConfig(string name, string Xvalue)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(Application.ExecutablePath + ".config");
            XmlNode node = doc.SelectSingleNode(@"//add[@key='" + name + "']");
            XmlElement ele = (XmlElement)node;
            ele.SetAttribute("value", Xvalue);
            doc.Save(Application.ExecutablePath + ".config");
        }


        public static string GetConfig(string key)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(Application.ExecutablePath + ".config");
            XmlNode node = doc.SelectSingleNode(@"//add[@key='" + key + "']");
            XmlElement ele = (XmlElement)node;
            string tmp = ele.GetAttribute("value");
            return tmp;
        }
    }
}
