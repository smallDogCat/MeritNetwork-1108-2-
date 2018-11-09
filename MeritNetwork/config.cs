using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace MeritNetwork
{
    public class config
    {
        public static void UpdateConfig(string filename, string name, string Xvalue)
        {
            int ss = 2;
            int s1 = 3;
            XmlDocument doc = new XmlDocument();
            doc.Load(filename + ".xml");
            XmlNode node = doc.SelectSingleNode(@"//add[@key='" + name + "']");
            XmlElement ele = (XmlElement)node;
            ele.SetAttribute("value", Xvalue);
            doc.Save(Application.ExecutablePath + ".config");
        }


        public static string GetConfig(string filename, string key)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(filename + ".xml");
            XmlNode node = doc.SelectSingleNode(@"//add[@key='" + key + "']");
            XmlElement ele = (XmlElement)node;
            string tmp = ele.GetAttribute("value");
            return tmp;
        }
    }
}
