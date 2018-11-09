using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.Configuration;
using System.Net;
using System.Xml;
using System.IO;
using System.Net.Sockets;
using Microsoft.Office.Interop.Excel;

namespace MeritNetwork
{
    public delegate void MYdelegate1(int md);
    public delegate double MYdelegate2(string md);
    public partial class Form1 : Form
    {
        private string name;
        private string _chan_num;
        private string _start_freq;
        private string _end_freq;
        List<string> Sfreq_list = new List<string>();
        List<int> listBox_num = new List<int>();
        List<string> listBox_chan = new List<string>();
        List<List<double>> amp_list = new List<List<double>>();
        List<List<double>> phase_list = new List<List<double>>();

        public static string control_addr;//控制模块地址
        public static string control_port;//控制模块端口
        public static string local_addr;//本机地址
        public static string local_port;//本机端口

        NetWorkAnalyser netWork = new NetWorkAnalyser();
        MessageForm message;

        public Form1()
        {
            InitializeComponent();
            Console.Write("gouzaohanshu----start");
            this.ModelSelect_groupControl.Visible = false;

            NChan_textEdit.Enabled = false;
            NStart_textEdit.Enabled = false;
            NEnd_textEdit.Enabled = false;
            NModel_simpleButton.Enabled = false;

            wave_checkEdit.Checked
                = loss_checkEdit.Checked
                = phase_checkEdit.Checked
                = Isolation_checkEdit.Checked
                = Uphase_checkEdit.Checked
                = Uamp_checkEdit.Checked
                = all_checkEdit.Checked;
            comboBox_single_CHAN.Enabled = false;
            comboBox_multi_CHAN.Enabled = false;

            _chan_num = config.GetConfig("BWPS20-3010-3350-20-1TR", "_Chan_num");
            _start_freq = config.GetConfig("BWPS20-3010-3350-20-1TR", "_Start_freq");
            _end_freq = config.GetConfig("BWPS20-3010-3350-20-1TR", "_End_freq");
            //Point_label.Text = "通道数：+'"_chan_num"'+\n工作频率：+'"_start_freq"'MHz~'"_end_freq"'MHz";
            Point_label.Text = "通道数：" + _chan_num + "\n工作频率：" + _start_freq + "MHz~" + _end_freq + "MHz";
        
            try
            {
                control_addr = Appconfig.GetConfig("CONTROLADDR");
                control_port = Appconfig.GetConfig("CONTROLPORT");
                local_addr = Appconfig.GetConfig("LOCALADDR");
                local_port = Appconfig.GetConfig("LOCALPORT");
                Console.Write("gouzaohanshu----try1");
                //于仪表建立连接
                netWork.NetWork_GPIB("17");
                Console.Write("gouzaohanshu----try2");
                //矢网：打开输出POWER
                netWork.NetWork_write("OUTP:STATe ON");//一般没用
            }
            catch (Exception)
            {
                //Console.Write("gouzaohanshu----trex");
                MessageBox.Show("仪表连接错误！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            for (int j = 0; j < int.Parse(_chan_num); j++)
            {
                comboBox_single_CHAN.Items.AddRange(new object[] { j + 1 });
                comboBox_multi_CHAN.Items.AddRange(new object[] { j + 1 });
            }

            int i = 0;
            XmlDocument mydoc = new XmlDocument();
            mydoc.Load("BWPS20-3010-3350-20-1TR.xml");
            XmlNodeList xndl = mydoc.SelectSingleNode("//appSettings").ChildNodes;
            foreach (var item in xndl)
            {
                XmlElement xmle = (XmlElement)item;
                if (xmle.GetAttribute("key") == "Sfreq")
                {
                    Sfreq_list.Add(xmle.GetAttribute("value"));
                    Freq_listBoxControl.Items.Add(Sfreq_list[i] + " MHz");
                    i++;
                }
            }

            Console.Write("gouzaohanshu----finish");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string str = Interaction.InputBox("请输入您的姓名：", "姓名输入", "", -1, -1);
            while (str == "")
            {
                MessageBox.Show("请输入姓名！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                str = Interaction.InputBox("请输入您的姓名：", "姓名输入", "", -1, -1);
            }
            name = str;
           
            //this.Analysis_spreadsheetControl.LoadDocument("tlp1.xlsx");
        }

        private void 打开ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (DialogResult.OK == this.openFileDialog1.ShowDialog())
            {
                this.Analysis_spreadsheetControl.LoadDocument(this.openFileDialog1.FileName);
            }
        }

        private void 保存ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string filename = null;
            string Pnumber = this.Pnum_textEdit.Text;
            if (Pnumber == "")
            {
                MessageBox.Show("请输入产品编号！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
            }
            else
            {
                this.Analysis_spreadsheetControl.Document.Worksheets[0].Cells["B"+1].Value = this.Pnum_textEdit.Text;
                if (BWPS1_radioButton.Checked == true)
                {
                    filename = "BWPS20-3010-3350-20-1TR";
                }
                else if (BWPS2_radioButton.Checked == true)
                {
                    filename = "BWPS20-3010-3350-20-1R";
                }
                else if (BWPS3_radioButton.Checked == true)
                {
                    filename = "BWPS20-15-125-20-1";
                }
                else if (BWPS4_radioButton.Checked == true)
                {
                    filename = "BWPS20-15-485-20-1";
                }
                else if (BWG_radioButton.Checked == true)
                {
                    filename = "BWG25235232-2.5-3.5";
                }
                else if (NModel_radioButton.Checked == true)
                {
                    filename = "B-None-Model";
                }

                DateTime curr = new DateTime();
                curr = DateTime.Now;
                this.saveFileDialog1.FileName = Pnumber + "_" + curr.Year + curr.Month + curr.Day + "_" + curr.Hour + "." + curr.Minute;
                if (!Directory.Exists("E:\\功分器测试数据\\" + filename + "\\" + curr.Year + curr.Month + curr.Day))
                {
                    Directory.CreateDirectory("E:\\功分器测试数据\\" + filename + "\\" + curr.Year + curr.Month + curr.Day);
                }
                this.saveFileDialog1.InitialDirectory = "E:\\功分器测试数据\\" + filename + "\\" + curr.Year + curr.Month + curr.Day;
                if (this.saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    this.Analysis_spreadsheetControl.SaveDocument(this.saveFileDialog1.FileName);
                }
                this.Pnum_textEdit.Text = "";
            }
        }

        private void Add_Button_Click(object sender, EventArgs e)
        {
            double _start = 0, _end = 0;
            string filename = null;
            if (BWPS1_radioButton.Checked == true)
            {
                _start = 3010;
                _end = 3350;
                filename = "BWPS20-3010-3350-20-1TR.xml";
            }
            else if (BWPS2_radioButton.Checked == true)
            {
                _start = 3010;
                _end = 3350;
                filename = "BWPS20-3010-3350-20-1R.xml";
            }
            else if (BWPS3_radioButton.Checked == true)
            {
                _start = 15;
                _end = 125;
                filename = "BWPS20-15-125-20-1.xml";
            }
            else if (BWPS4_radioButton.Checked == true)
            {
                _start = 15;
                _end = 485;
                filename = "BWPS20-15-485-20-1.xml";
            }
            else if (BWG_radioButton.Checked == true)
            {
                _start = 2500;
                _end = 3500;
                filename = "BWG25235232-2.5-3.5.xml";
            }
            else if (NModel_radioButton.Checked == true)
            {
                _start = double.Parse(this.NStart_textEdit.Text);
                _end = double.Parse(this.NEnd_textEdit.Text);
                filename = "B-None-Model.xml";
            }
            if (_start <= double.Parse(this.Freq_textEdit.Text) && double.Parse(this.Freq_textEdit.Text) <= _end)
            {
                if (!Freq_listBoxControl.Items.Contains(this.Freq_textEdit.Text + " MHz"))
                {
                    Freq_listBoxControl.Items.Add(Freq_textEdit.Text + " MHz");
                    Sfreq_list.Add(Freq_textEdit.Text);

                    XmlDocument mydoc = new XmlDocument();
                    mydoc.Load(filename);
                    XmlNode xnd = mydoc.SelectSingleNode("//appSettings");
                    XmlElement xle = mydoc.CreateElement("add");
                    xle.SetAttribute("key", "Sfreq");
                    xle.SetAttribute("value", Freq_textEdit.Text);
                    xnd.AppendChild(xle);
                    mydoc.Save(filename);
                }
            }
            else
            {
                MessageBox.Show("请输入频率范围频点！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            }

            Freq_textEdit.Text = "";
        }

        private void Clean_Button_Click(object sender, EventArgs e)
        {
            DialogResult r1 = MessageBox.Show("确认清空频点？清空后所有频点将被删除，下次使用时将要全部重新设置！", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if ((int)r1 == 1)
            {
                Freq_listBoxControl.Items.Clear();
                Sfreq_list.Clear();
                XmlDocument mydoc = new XmlDocument();
                string filename = null;
                if (BWPS1_radioButton.Checked == true)
                {
                    filename = "BWPS20-3010-3350-20-1TR.xml";
                }
                else if (BWPS2_radioButton.Checked == true)
                {
                    filename = "BWPS20-3010-3350-20-1R.xml";
                }
                else if (BWPS3_radioButton.Checked == true)
                {
                    filename = "BWPS20-15-125-20-1.xml";
                }
                else if (BWPS4_radioButton.Checked == true)
                {
                    filename = "BWPS20-15-485-20-1.xml";
                }
                else if (BWG_radioButton.Checked == true)
                {
                    filename = "BWG25235232-2.5-3.5.xml";
                }
                else if (NModel_radioButton.Checked == true)
                {
                    filename = "B-None-Model.xml";
                }
                mydoc.Load(filename);
                //XmlNodeList xndl = mydoc.SelectSingleNode("//appSettings").ChildNodes;
                //foreach (var item in xndl)
                //{
                //    XmlElement xle = (XmlElement)item;
                //    if (xle.GetAttribute("key") == "Sfreq")
                //    {
                //        //xle.RemoveAll();
                //        xle.ParentNode.RemoveChild(xle);
                //    }
                //}
                XmlNodeList xndl = mydoc.SelectNodes("//add");
                foreach (var item in xndl)
                {
                    XmlElement xle = (XmlElement)item;
                    if (xle.GetAttribute("key") == "Sfreq")
                    {
                        //xle.RemoveAll();
                        xle.ParentNode.RemoveChild(xle);
                    }
                }
                mydoc.Save(filename);
            }
        }

        private void BWPS2_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            Sfreq_list.Clear();
            Freq_listBoxControl.Items.Clear();
            _chan_num = config.GetConfig("BWPS20-3010-3350-20-1R", "_Chan_num");
            _start_freq = config.GetConfig("BWPS20-3010-3350-20-1R", "_Start_freq");
            _end_freq = config.GetConfig("BWPS20-3010-3350-20-1R", "_End_freq");
            Point_label.Text = "通道数：" + _chan_num + "\n工作频率：" + _start_freq + "MHz~" + _end_freq + "MHz";

            comboBox_single_CHAN.Items.Clear();
            comboBox_multi_CHAN.Items.Clear();
            for (int j = 0; j < int.Parse(_chan_num); j++)
            {
                comboBox_single_CHAN.Items.AddRange(new object[] { j + 1 });
                comboBox_multi_CHAN.Items.AddRange(new object[] { j + 1 });
            }

            int i = 0;
            XmlDocument mydoc = new XmlDocument();
            mydoc.Load("BWPS20-3010-3350-20-1R.xml");
            XmlNodeList xndl = mydoc.SelectSingleNode("//appSettings").ChildNodes;
            foreach (var item in xndl)
            {
                XmlElement xmle = (XmlElement)item;
                if (xmle.GetAttribute("key") == "Sfreq")
                {
                    Sfreq_list.Add(xmle.GetAttribute("value"));
                    Freq_listBoxControl.Items.Add(Sfreq_list[i] + " MHz");
                    i++;
                }
            }
        }

        private void BWPS3_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            Sfreq_list.Clear();
            Freq_listBoxControl.Items.Clear();
            _chan_num = config.GetConfig("BWPS20-15-125-20-1", "_Chan_num");
            _start_freq = config.GetConfig("BWPS20-15-125-20-1", "_Start_freq");
            _end_freq = config.GetConfig("BWPS20-15-125-20-1", "_End_freq");
            Point_label.Text = "通道数：" + _chan_num + "\n工作频率：" + _start_freq + "MHz~" + _end_freq + "MHz";

            comboBox_single_CHAN.Items.Clear();
            comboBox_multi_CHAN.Items.Clear();
            for (int j = 0; j < int.Parse(_chan_num); j++)
            {
                comboBox_single_CHAN.Items.AddRange(new object[] { j + 1 });
                comboBox_multi_CHAN.Items.AddRange(new object[] { j + 1 });
            }

            int i = 0;
            XmlDocument mydoc = new XmlDocument();
            mydoc.Load("BWPS20-15-125-20-1.xml");
            XmlNodeList xndl = mydoc.SelectSingleNode("//appSettings").ChildNodes;
            foreach (var item in xndl)
            {
                XmlElement xmle = (XmlElement)item;
                if (xmle.GetAttribute("key") == "Sfreq")
                {
                    Sfreq_list.Add(xmle.GetAttribute("value"));
                    Freq_listBoxControl.Items.Add(Sfreq_list[i] + " MHz");
                    i++;
                }
            }
        }

        private void BWPS4_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            Sfreq_list.Clear();
            Freq_listBoxControl.Items.Clear();
            _chan_num = config.GetConfig("BWPS20-15-485-20-1", "_Chan_num");
            _start_freq = config.GetConfig("BWPS20-15-485-20-1", "_Start_freq");
            _end_freq = config.GetConfig("BWPS20-15-485-20-1", "_End_freq");
            Point_label.Text = "通道数：" + _chan_num + "\n工作频率：" + _start_freq + "MHz~" + _end_freq + "MHz";

            comboBox_single_CHAN.Items.Clear();
            comboBox_multi_CHAN.Items.Clear();
            for (int j = 0; j < int.Parse(_chan_num); j++)
            {
                comboBox_single_CHAN.Items.AddRange(new object[] { j + 1 });
                comboBox_multi_CHAN.Items.AddRange(new object[] { j + 1 });
            }

            int i = 0;
            XmlDocument mydoc = new XmlDocument();
            mydoc.Load("BWPS20-15-485-20-1.xml");
            XmlNodeList xndl = mydoc.SelectSingleNode("//appSettings").ChildNodes;
            foreach (var item in xndl)
            {
                XmlElement xmle = (XmlElement)item;
                if (xmle.GetAttribute("key") == "Sfreq")
                {
                    Sfreq_list.Add(xmle.GetAttribute("value"));
                    Freq_listBoxControl.Items.Add(Sfreq_list[i] + " MHz");
                    i++;
                }
            }
        }

        private void BWG_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            Sfreq_list.Clear();
            Freq_listBoxControl.Items.Clear();
            _chan_num = config.GetConfig("BWG25235232-2.5-3.5", "_Chan_num");
            _start_freq = config.GetConfig("BWG25235232-2.5-3.5", "_Start_freq");
            _end_freq = config.GetConfig("BWG25235232-2.5-3.5", "_End_freq");
            Point_label.Text = "通道数：" + _chan_num + "\n工作频率：" + _start_freq + "MHz~" + _end_freq + "MHz";

            comboBox_single_CHAN.Items.Clear();
            comboBox_multi_CHAN.Items.Clear();
            for (int j = 0; j < int.Parse(_chan_num); j++)
            {
                comboBox_single_CHAN.Items.AddRange(new object[] { j + 1 });
                comboBox_multi_CHAN.Items.AddRange(new object[] { j + 1 });
            }

            int i = 0;
            XmlDocument mydoc = new XmlDocument();
            mydoc.Load("BWG25235232-2.5-3.5.xml");
            XmlNodeList xndl = mydoc.SelectSingleNode("//appSettings").ChildNodes;
            foreach (var item in xndl)
            {
                XmlElement xmle = (XmlElement)item;
                if (xmle.GetAttribute("key") == "Sfreq")
                {
                    Sfreq_list.Add(xmle.GetAttribute("value"));
                    Freq_listBoxControl.Items.Add(Sfreq_list[i] + " MHz");
                    i++;
                }
            }
        }

        private void NModel_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            //MessageBox.Show();
            if (NModel_radioButton.Checked == false)
            {
                NChan_textEdit.Enabled = false;
                NStart_textEdit.Enabled = false;
                NEnd_textEdit.Enabled = false;
                NModel_simpleButton.Enabled = false;
            }
            else
            {
                NChan_textEdit.Enabled = true;
                NStart_textEdit.Enabled = true;
                NEnd_textEdit.Enabled = true;
                NModel_simpleButton.Enabled = true;
            }
            Sfreq_list.Clear();
            Freq_listBoxControl.Items.Clear();
            _chan_num = config.GetConfig("B-None-Model", "_Chan_num");
            _start_freq = config.GetConfig("B-None-Model", "_Start_freq");
            _end_freq = config.GetConfig("B-None-Model", "_End_freq");
            Point_label.Text = "通道数：" + _chan_num + "\n工作频率：" + _start_freq + "MHz~" + _end_freq + "MHz";

            comboBox_single_CHAN.Items.Clear();
            comboBox_multi_CHAN.Items.Clear();
            for (int j = 0; j < int.Parse(_chan_num); j++)
            {
                comboBox_single_CHAN.Items.AddRange(new object[] { j + 1 });
                comboBox_multi_CHAN.Items.AddRange(new object[] { j + 1 });
            }

            NChan_textEdit.Text = _chan_num;
            NStart_textEdit.Text = _start_freq;
            NEnd_textEdit.Text = _end_freq;

            int i = 0;
            XmlDocument mydoc = new XmlDocument();
            mydoc.Load("B-None-Model.xml");
            XmlNodeList xndl = mydoc.SelectSingleNode("//appSettings").ChildNodes;
            foreach (var item in xndl)
            {
                XmlElement xmle = (XmlElement)item;
                if (xmle.GetAttribute("key") == "Sfreq")
                {
                    Sfreq_list.Add(xmle.GetAttribute("value"));
                    Freq_listBoxControl.Items.Add(Sfreq_list[i] + " MHz");
                    i++;
                }
            }
        }
 
        private void NModel_simpleButton_Click(object sender, EventArgs e)
        {
            XmlDocument mydoc = new XmlDocument();
            mydoc.Load("B-None-Model.xml");
            XmlNodeList xndl = mydoc.SelectSingleNode("//appSettings").ChildNodes;
            XmlNode xnd = mydoc.SelectSingleNode("//add[@key='Sfreq']");
            if (xnd == null)
            {
                XmlNode xnd1 = mydoc.SelectSingleNode("//add[@key='_Chan_num']");
                XmlElement xle1 = (XmlElement)xnd1;
                if (xle1.GetAttribute("value") != NChan_textEdit.Text)
                {
                    xle1.SetAttribute("value", NChan_textEdit.Text);
                }

                XmlNode xnd2 = mydoc.SelectSingleNode("//add[@key='_Start_freq']");
                XmlElement xle2 = (XmlElement)xnd2;
                if (xle2.GetAttribute("value") != NStart_textEdit.Text)
                {
                    xle2.SetAttribute("value", NStart_textEdit.Text);
                }

                XmlNode xnd3 = mydoc.SelectSingleNode("//add[@key='_End_freq']");
                XmlElement xle3 = (XmlElement)xnd3;
                if (xle3.GetAttribute("value") != NEnd_textEdit.Text)
                {
                    xle3.SetAttribute("value", NEnd_textEdit.Text);
                }

                mydoc.Save("B-None-Model.xml");
            }
            else
            {
                XmlElement xle = (XmlElement)xnd;
                double small = double.Parse(xle.GetAttribute("value"));
                double big = double.Parse(xle.GetAttribute("value"));
                foreach (var item in xndl)
                {
                    XmlElement xel = (XmlElement)item;
                    if (xel.GetAttribute("key") == "Sfreq")
                    {
                        if (double.Parse(xel.GetAttribute("value")) < small)
                            small = double.Parse(xel.GetAttribute("value"));
                        if (double.Parse(xel.GetAttribute("value")) > big)
                            big = double.Parse(xel.GetAttribute("value"));
                    }
                }
                if (0 > double.Parse(this.NChan_textEdit.Text) || 32 < double.Parse(this.NChan_textEdit.Text))
                {
                    MessageBox.Show("通道数为0~32！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                    this.NChan_textEdit.Text = "";
                }
                if (double.Parse(this.NStart_textEdit.Text) > double.Parse(this.NEnd_textEdit.Text))
                {
                    MessageBox.Show("频率范围有误，请检查！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                }

                else if (double.Parse(this.NStart_textEdit.Text) > small || double.Parse(this.NEnd_textEdit.Text) < big)
                {
                    MessageBox.Show("频率范围与所选频点有冲突，若想修改频率范围，请先删除冲突频点！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                    if (double.Parse(this.NStart_textEdit.Text) > small)
                    {
                        this.NStart_textEdit.Text = "";

                    }
                    else
                    {
                        this.NEnd_textEdit.Text = "";
                    }
                }
                else
                {
                    XmlNode xnd1 = mydoc.SelectSingleNode("//add[@key='_Chan_num']");
                    XmlElement xle1 = (XmlElement)xnd1;
                    if (xle1.GetAttribute("value") != NChan_textEdit.Text)
                    {
                        xle1.SetAttribute("value", NChan_textEdit.Text);
                    }

                    XmlNode xnd2 = mydoc.SelectSingleNode("//add[@key='_Start_freq']");
                    XmlElement xle2 = (XmlElement)xnd2;
                    if (xle2.GetAttribute("value") != NStart_textEdit.Text)
                    {
                        xle2.SetAttribute("value", NStart_textEdit.Text);
                    }

                    XmlNode xnd3 = mydoc.SelectSingleNode("//add[@key='_End_freq']");
                    XmlElement xle3 = (XmlElement)xnd3;
                    if (xle3.GetAttribute("value") != NEnd_textEdit.Text)
                    {
                        xle3.SetAttribute("value", NEnd_textEdit.Text);
                    }

                    mydoc.Save("B-None-Model.xml");
                }
            }
        }

        private void 删除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult r1 = MessageBox.Show("确认删除？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
            if ((int)r1 == 1)
            {
                for (int i = 0; i < Sfreq_list.Count; i++)
                {
                    if ((Sfreq_list[i] + " MHz") == Freq_listBoxControl.SelectedItem.ToString())
                    {
                        Sfreq_list.RemoveAt(i);
                    }
                }

                XmlDocument mydoc = new XmlDocument();
                string filename = null;
                if (BWPS1_radioButton.Checked == true)
                {
                    filename = "BWPS20-3010-3350-20-1TR.xml";
                }
                else if (BWPS2_radioButton.Checked == true)
                {
                    filename = "BWPS20-3010-3350-20-1R.xml";
                }
                else if (BWPS3_radioButton.Checked == true)
                {
                    filename = "BWPS20-15-125-20-1.xml";
                }
                else if (BWPS4_radioButton.Checked == true)
                {
                    filename = "BWPS20-15-485-20-1.xml";
                }
                else if (BWG_radioButton.Checked == true)
                {
                    filename = "BWG25235232-2.5-3.5.xml";
                }
                else if (NModel_radioButton.Checked == true)
                {
                    filename = "B-None-Model.xml";
                }
                mydoc.Load(filename);
                XmlNodeList xndl = mydoc.SelectNodes("//add[@key='Sfreq']");
                foreach (var item in xndl)
                {
                    XmlElement xel = (XmlElement)item;
                    if ((xel.GetAttribute("value") + " MHz") == Freq_listBoxControl.SelectedItem.ToString())
                    {
                        xel.ParentNode.RemoveChild(xel);
                    }
                }
                mydoc.Save(filename);
                Freq_listBoxControl.Items.RemoveAt(Freq_listBoxControl.SelectedIndex);
            }
        }

        private void Start_simpleButton_Click(object sender, EventArgs e)
        {
            string filename = null;
            if (BWPS1_radioButton.Checked == true)
            {
                filename = "BWPS20-3010-3350-20-1";
            }
            else if (BWPS2_radioButton.Checked == true)
            {
                filename = "BWPS20-3010-3350-20-1";//BWPS1与BWPS2共用一组校准数据
            }
            else if (BWPS3_radioButton.Checked == true)
            {
                filename = "BWPS20-15-125-20-1";
            }
            else if (BWPS4_radioButton.Checked == true)
            {
                filename = "BWPS20-15-485-20-1";
            }
            else if (BWG_radioButton.Checked == true)
            {
                filename = "BWG25235232-2.5-3.5";
            }
            else if (NModel_radioButton.Checked == true)
            {
                filename = "B-None-Model";
            }

            if (!Directory.Exists("E:\\功分器校准数据\\" + filename))
            {
                Directory.CreateDirectory("E:\\功分器校准数据\\" + filename);
            }

            if (!File.Exists("E:\\功分器校准数据\\" + filename + "\\" + filename + ".xlsx"))
            {
                MessageBox.Show("该型号产品校准数据缺失，请补上！校准文件请放在" + "E:\\功分器校准数据\\" + filename, "提示", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
            }
            else
            {
                if (NModel_radioButton.Checked == true)
                {
                    DialogResult r0 = MessageBox.Show("当前是非型号产品，校准文件是否已替换为测试产品校准文件？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                    if ((int)r0 == 6)
                    {
                        DialogResult r1 = MessageBox.Show("被测件是否和开关矩阵按顺序连接？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                        if ((int)r1 == 6)
                        {
                            if (!this.bgwork_test1.IsBusy)
                                this.bgwork_test1.RunWorkerAsync();
                            else
                                MessageBox.Show("程序正在运行。。。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else
                {
                    DialogResult r1 = MessageBox.Show("被测件是否和开关矩阵按顺序连接？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                    if ((int)r1 == 6)
                    {
                        if (!this.bgwork_test1.IsBusy)
                            this.bgwork_test1.RunWorkerAsync();
                        else
                            MessageBox.Show("程序正在运行。。。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void End_simpleButton_Click(object sender, EventArgs e)
        {
            this.bgwork_test1.CancelAsync();
        }

        private void bgwork_test1_DoWork(object sender, DoWorkEventArgs e)
        {
            amp_list.Clear();
            phase_list.Clear();

            string _chan_num = "";
            string filename = null;
            double wave = 0, loss = 0;

            if (BWPS1_radioButton.Checked == true)
            {
                filename = "BWPS20-3010-3350-20-1TR.xml";
                _chan_num = config.GetConfig("BWPS20-3010-3350-20-1TR", "_Chan_num");
            }
            else if (BWPS2_radioButton.Checked == true)
            {
                filename = "BWPS20-3010-3350-20-1R.xml";
                _chan_num = config.GetConfig("BWPS20-3010-3350-20-1R", "_Chan_num");
            }
            else if (BWPS3_radioButton.Checked == true)
            {
                filename = "BWPS20-15-125-20-1.xml";
                _chan_num = config.GetConfig("BWPS20-15-125-20-1", "_Chan_num");
            }
            else if (BWPS4_radioButton.Checked == true)
            {
                filename = "BWPS20-15-485-20-1.xml";
                _chan_num = config.GetConfig("BWPS20-15-485-20-1", "_Chan_num");
            }
            else if (BWG_radioButton.Checked == true)
            {
                filename = "BWG25235232-2.5-3.5.xml";
                _chan_num = config.GetConfig("BWG25235232-2.5-3.5", "_Chan_num");
            }
            else if (NModel_radioButton.Checked == true)
            {
                filename = "B-None-Model.xml";
                _chan_num = config.GetConfig("B-None-Model", "_Chan_num");
            }

            XmlDocument mydoc = new XmlDocument();
            mydoc.Load(filename);
            XmlNodeList xndl = mydoc.SelectNodes("//add[@key='Sfreq']");
            int length = xndl.Count;
            this.splashScreenManager1.ShowWaitForm();
            this.splashScreenManager1.SetWaitFormCaption("正在进行表格的设计与加载，请稍等~~~");
            Create_excel(length);      
            this.Analysis_spreadsheetControl.Invoke(new EventHandler(delegate {
                this.Analysis_spreadsheetControl.LoadDocument("tlp1.xlsx");
            }));
            this.Analysis_spreadsheetControl.Invoke(new EventHandler(delegate { 
                this.Analysis_spreadsheetControl.Document.Worksheets[0].Cells["B" + 2].Value = name; 
            }));
            
            this.splashScreenManager1.CloseWaitForm();
            if (Single_Chan_radioButton.Checked == true)
            {
                int chan = int.Parse(comboBox_single_CHAN.Text);
                SendCmd_Parallel(chan);
                double[] phase = new double[1];
                double[] uphase = new double[1];
                double[] uamp = new double[1];
                if (this.wave_checkEdit.Checked == true)
                {
                    wave = Wave(filename);
                }
                if (this.loss_checkEdit.Checked == true)
                {
                    loss = Loss(filename);
                }
                if (this.phase_checkEdit.Checked == true)
                {
                    phase = Phase(filename);
                }
                if (this.Uphase_checkEdit.Checked == true)
                {
                    uphase = Uphase(phase_list);
                }
                if (this.Uamp_checkEdit.Checked == true)
                {
                    uamp = Uamp(amp_list);
                }
                if (Isolation_checkEdit.Checked == true)
                {
                    Isolation(filename, int.Parse(_chan_num) / 2);
                }
            }
            else if (Multi_Chan_radioButton.Checked == true)
            {
                int len = listBox_chan.Count();
                double[] phase = new double[len];
                double[] uphase = new double[len];
                double[] uamp = new double[len];
                int chan = 0;
                for (int i = 0; i < len; i++)
                {
                    if (bgwork_test1.CancellationPending)
                    {
                        e.Cancel = true; //这里才真正取消  
                        break;
                    }
                    chan = int.Parse(listBox_chan[i]);
                    SendCmd_Parallel(chan);
                    if (this.wave_checkEdit.Checked == true)
                    {
                        wave = Wave(filename);
                    }
                    if (this.loss_checkEdit.Checked == true)
                    {
                        loss = Loss(filename);

                    }
                    if (this.phase_checkEdit.Checked == true)
                    {
                        phase = Phase(filename);
                    }
                }
                if (this.Uphase_checkEdit.Checked == true)
                {
                    uphase = Uphase(phase_list);
                }
                if (this.Uamp_checkEdit.Checked == true)
                {
                    uamp = Uamp(amp_list);
                }
                if (Isolation_checkEdit.Checked == true)
                {
                    Isolation(filename, int.Parse(_chan_num) / 2);
                }
            }
            else if (All_Chan_radioButton.Checked == true)
            {
                int len = int.Parse(_chan_num);
                double[] phase = new double[len];
                double[] uphase = new double[len];
                double[] uamp = new double[len];
                int chan = 0;
                for (int i = 0; i < len; i++)
                {
                    if (bgwork_test1.CancellationPending)
                    {
                        e.Cancel = true; //这里才真正取消  
                        break;
                    }
                    chan = i + 1;
                    SendCmd_Parallel(chan);
                    if (this.wave_checkEdit.Checked == true)
                    {
                        wave = Wave(filename);
                    }
                    if (this.loss_checkEdit.Checked == true)
                    {
                        loss = Loss(filename);

                    }
                    if (this.phase_checkEdit.Checked == true)
                    {
                        phase = Phase(filename);
                    }
                }
                if (this.Uphase_checkEdit.Checked == true)
                {
                    uphase = Uphase(phase_list);
                }
                if (this.Uamp_checkEdit.Checked == true)
                {
                    uamp = Uamp(amp_list);
                }
                if (Isolation_checkEdit.Checked == true)
                {
                    Isolation(filename, int.Parse(_chan_num) / 2);
                }
            }
        }
        private double Wave(string file) //端口驻波
        {
            double[] markarr = sw_mark(file);
            double[] s1 = sw_data("CH1_S11_1", markarr);
            double s11 = s1.Max();
            //isLoad2000 = false;
            double[] s2 = sw_data("CH1_S11_4", markarr);
            double s22 = s2.Max();
            double maxp;//驻波为s11和s22中的较大值
            if (s11 > s22)
                maxp = s11;
            else
                maxp = s22;
            return maxp;
        }
        private double Loss(string file)
        {
            double[] markarr = sw_mark(file);
            double[] s1 = sw_data("CH1_S11_2", markarr);//幅度
            List<double> res = new List<double>();
            for (int i = 0; i < s1.Length; i++)
            {
                res.Add(Math.Abs(s1[i]));
            }
            //double maxp = res.Max();       //损耗 
            //amp_list.Add(res);
            //return maxp;
            return 1;
        }
        private double[] Phase(string file)
        {
            double[] markarr = sw_mark(file);
            double[] s1 = sw_data("CH1_S11_3", markarr);//相位
            List<double> res = new List<double>();
            for (int i = 0; i < s1.Length; i++)
            {
                res.Add(s1[i]);
            }
            phase_list.Add(res);
            return s1;
        }
        public static string fileM;
        private void Isolation(string file, int num)
        {
            MYdelegate1 mydele1 = new MYdelegate1(SendCmd_Parallel);
            MYdelegate2 mydele2 = new MYdelegate2(Loss);

            fileM = file;
            double[] phaseI = new double[num];
            message = new MessageForm(mydele1,mydele2);
            //弹出窗口
            message.ShowDialog();
            while (message.DialogResult == DialogResult.OK)
            {
                message = new MessageForm(mydele1,mydele2);
                message.ShowDialog();
            }

        }
        private double[] Uphase(List<List<double>> list)
        {
            int len = list.Count;//通道数目
            int len1 = list[0].Count;//频点数目
            double[] uphase = new double[len1];
            for (int j = 0; j < len1; j++)
            {
                double max = list[0][j];
                double min = list[0][j];
                for (int i = 0; i < len; i++)
                {
                    if (list[i][j] > max)
                    {
                        max = list[i][j];
                    }
                    if (list[i][j] < min)
                    {
                        min = list[i][j];
                    }
                }
                uphase[j] = max - min;
            }
            return uphase;

        }
        private double[] Uamp(List<List<double>> list)
        {
            int len = list.Count;//通道数目
            int len1 = list[0].Count;//频点数目
            double[] uamp = new double[len1];
            for (int j = 0; j < len1; j++)
            {
                double max = list[0][j];
                double min = list[0][j];
                for (int i = 0; i < len; i++)
                {
                    if (list[i][j] > max)
                    {
                        max = list[i][j];
                    }
                    if (list[i][j] < min)
                    {
                        min = list[i][j];
                    }
                }
                uamp[j] = max - min;
            }
            return uamp;
        }
        private double[] sw_mark(string filename)
        {
            int i = 0;
            //设置mark点
            XmlDocument mydoc = new XmlDocument();
            mydoc.Load(filename);
            XmlNodeList xndl = mydoc.SelectNodes("//add[@key='Sfreq']");
            int len = xndl.Count;
            double[] mark = new double[len];
            foreach (var item in xndl)
            {
                XmlElement xle = (XmlElement)item;
                mark[i] = double.Parse(xle.GetAttribute("value"));
                netWork.NetWork_write("CALC:MARK" + (++i) + ":X " + mark[i] + "MHz");
                i++;
            }
            return mark;
        }
        private double[] sw_data(string tblocstr, double[] mark)
        {
            object res = "";

            int len = mark.Length;
            double[] restmp = new double[len];

            //矢网：选择测试窗口，测试轨迹
            netWork.NetWork_write("CALCulate:PARameter:SELect '" + tblocstr + "'");
            //矢网：打开MARK点
            netWork.NetWork_write("calculate:marker on");//一般没用

            for (int i = 0; i < len; i++)
            {
                restmp[i] = netWork.NetWork_read("CALC:MARK" + (++i) + ":Y?");
                //CALC:MARK1:Y?
            }
            //double res1 = restmp.Max();

            //关闭连接
            netWork.NetWork_Close();
            return restmp;
        }

        private void 型号ToolStripMenuItem_MouseEnter(object sender, EventArgs e)
        {
            this.ModelSelect_groupControl.Visible = true;
        }

        private void 型号ToolStripMenuItem_MouseLeave(object sender, EventArgs e)
        {
            //this.ModelSelect_groupControl.Visible = false;
        }

        private void ModelSelect_groupControl_MouseLeave(object sender, EventArgs e)
        {
            if (!this.ModelSelect_groupControl.ClientRectangle.Contains(this.ModelSelect_groupControl.PointToClient(Control.MousePosition)))
            {
                //base.OnMouseLeave(e);
                ModelSelect_groupControl.Visible = false;
                //ModelSelect_groupControl.

            }
        }

        private void all_checkEdit_CheckedChanged(object sender, EventArgs e)
        {
            wave_checkEdit.Checked
                = loss_checkEdit.Checked
                = phase_checkEdit.Checked
                = Isolation_checkEdit.Checked
                = Uphase_checkEdit.Checked
                = Uamp_checkEdit.Checked
                = all_checkEdit.Checked;
        }

        private void Single_Chan_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (Single_Chan_radioButton.Checked == false)
            {
                comboBox_single_CHAN.Enabled = false;
            }
            else if (Single_Chan_radioButton.Checked == true)
            {
                comboBox_single_CHAN.Enabled = true;
            }
            if (Chan_listBox.Items.Count != 0)
            {
                Chan_listBox.Items.Clear();
                listBox_num.Clear();
                listBox_chan.Clear();
            }
        }

        private void All_Chan_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (All_Chan_radioButton.Checked == false)
            {
                Chan_all_checkBox.Checked = false;
            }
            else if (All_Chan_radioButton.Checked == true)
            {
                Chan_all_checkBox.Checked = true;
            }
            if (Chan_listBox.Items.Count != 0)
            {
                Chan_listBox.Items.Clear();
                listBox_num.Clear();
                listBox_chan.Clear();
            }
        }

        private void Multi_Chan_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (Multi_Chan_radioButton.Checked == false)
            {
                comboBox_multi_CHAN.Enabled = false;
                Chan_listBox.Enabled = false;
            }
            else if (Multi_Chan_radioButton.Checked == true)
            {
                comboBox_multi_CHAN.Enabled = true;
                Chan_listBox.Enabled = true;
            }
        }

        private void Chan_all_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (Chan_all_checkBox.Checked == false)
            {
                All_Chan_radioButton.Checked = false;
            }
            else if (Chan_all_checkBox.Checked == true)
            {
                All_Chan_radioButton.Checked = true;
            }
        }

        private void comboBox_multi_CHAN_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!Chan_listBox.Items.Contains(comboBox_multi_CHAN.Text))
            {
                Chan_listBox.Items.Add(comboBox_multi_CHAN.Text);
                listBox_num.Add(comboBox_multi_CHAN.SelectedIndex);
                listBox_chan.Add(comboBox_multi_CHAN.Text);
            }
        }

        private void Clean_Click(object sender, EventArgs e)
        {
            Chan_listBox.Items.Clear();
            listBox_num.Clear();
            listBox_chan.Clear();
        }

        private void 删除ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            int n = listBox_chan.Count();
            try
            {
                for (int i = 0; i < n; i++)
                {
                    if (listBox_chan[i] == Chan_listBox.SelectedItem.ToString())
                    {
                        listBox_num.RemoveAt(i);
                        listBox_chan.RemoveAt(i);
                        break;
                    }
                }
                Chan_listBox.Items.RemoveAt(Chan_listBox.SelectedIndex);
            }
            catch { }
        }

        private UdpClient udpc;
        bool UDP_IsudpcStart = false;                                                   //udp是否开启
        //bool winext = false;                                                            //程序退出

        string COMMAND = "00 00";                                                       //2BYTE         命令编号
        string COMMAND_END = "DD";

        private void SendCmd_Parallel(int chan)//并行数据的控制
        {
            string s4 = null;
            string s81 = null;
            string s82 = null;
            string s83 = null;
            string s84 = null;
            //数据发送控制指令
            COMMAND = "FF AA";//

            if (1 <= chan && chan <= 8)
            {
                s4 = "01";
                s81 = Convert.ToString(chan, 16);
                s82 = "00";
                s83 = "00";
                s84 = "00";
            }
            if (9 <= chan && chan <= 16)
            {
                s4 = "02";
                s81 = "00";
                s82 = Convert.ToString(chan - 8, 16);
                s83 = "00";
                s84 = "00";
            }
            if (17 <= chan && chan <= 24)
            {
                s4 = "03";
                s81 = "00";
                s82 = "00";
                s83 = Convert.ToString(chan - 16, 16);
                s84 = "00";
            }
            if (25 <= chan && chan <= 32)
            {
                s4 = "04";
                s81 = "00";
                s82 = "00";
                s83 = "00";
                s84 = Convert.ToString(chan - 24, 16);
            }

            UDP_Send(control_addr, Convert.ToInt32(control_port), COMMAND
                                + " " + s4 + " " + s81 + " " + s82 + " " + s83 + " " + s84);
        }

        private void UDP_Send(string ipaddr, int pnum, string msg)
        {
            try
            {
                if (UDP_IsudpcStart == false)
                {
                    IPEndPoint localip = new IPEndPoint(IPAddress.Parse(local_addr), Convert.ToInt32(local_port));
                    udpc = new UdpClient(localip);
                    UDP_IsudpcStart = true;
                }
                string smsg = "";

                smsg = msg + " " + COMMAND_END;         //整个数据包的构成
                string[] smsgs = smsg.Split(' ');
                byte[] sendbytes = new byte[smsgs.Length];
                for (int i = 0; i < smsgs.Length; i++)
                {
                    string send = Convert.ToString(Convert.ToInt32(smsgs[i], 16));//Convert.ToInt32(smsgs[i], 16),将16进制形式的smsgs[i]转换成十进制
                    sendbytes[i] = Convert.ToByte(send);
                }

                IPEndPoint remoteIP = new IPEndPoint(IPAddress.Parse(ipaddr), pnum);
                udpc.Send(sendbytes, sendbytes.Length, remoteIP);
            }
            catch (Exception e)
            {
                MessageBox.Show(null, e.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void Create_excel(int leng)
        {
            object MisValue = Type.Missing;
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();          
            Workbooks xlsbooks = app.Workbooks;
            //Workbook xlsbook = xlsbooks.Add(System.Windows.Forms.Application.StartupPath+@"\tlp.xlsx");//若用add方法，则后续不能用save，只能用saveas

            //第三个参数为true，只读；为false，可读写
            Workbook xlsbook = xlsbooks.Open(System.Windows.Forms.Application.StartupPath + @"\tlp.xlsx", MisValue, false, MisValue, MisValue, MisValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, MisValue, MisValue, MisValue, MisValue, MisValue);
            Worksheet xlsSheet = xlsbook.Worksheets[1];

            //添加行
            //Range xlsRow = (Range)xlsSheet.Rows[3, MisValue];
            //xlsRow.Insert(XlInsertShiftDirection.xlShiftDown, MisValue);

            //幅度
            for (int i = 0; i < leng; i++) 
            {
                Range xlsColumns = (Range)xlsSheet.Columns[4, MisValue];
                xlsColumns.Insert(XlInsertShiftDirection.xlShiftToRight, MisValue);
                Range combine = xlsSheet.Range[xlsSheet.Cells[3, 3], xlsSheet.Cells[3, 4]];
                combine.Merge();
            }
            for (int i = 0; i <= leng; i++)
            {
                if (i != leng)
                {
                    xlsSheet.Cells[4, 3 + i] = Sfreq_list[i] + "MHz";
                }
                else
                {
                    xlsSheet.Cells[4, 3 + i] = "插入损耗"; //插入损耗
                }
            }
            //相位
            for (int i = leng + 1; i < leng + 3; i++)
            {
                Range xlsColumns = (Range)xlsSheet.Columns[4+(leng+1), MisValue];
                xlsColumns.Insert(XlInsertShiftDirection.xlShiftToRight, MisValue);
                Range combine = xlsSheet.Range[xlsSheet.Cells[3, 4 + leng], xlsSheet.Cells[3, 4 + (leng + 1)]];
                combine.Merge();
            }
            for (int i = leng + 1; i <= leng + 3; i++)
            {

                xlsSheet.Cells[4, 3 + i] = Sfreq_list[i - (leng + 1)] + "MHz";
            }
            //隔离度
            for (int i = leng + 4; i < leng + 6; i++)
            {
                Range xlsColumns = (Range)xlsSheet.Columns[4 + (leng + 4), MisValue];
                xlsColumns.Insert(XlInsertShiftDirection.xlShiftToRight, MisValue);
                Range combine = xlsSheet.Range[xlsSheet.Cells[3, 4 + (leng+3)], xlsSheet.Cells[3, 4 + (leng + 4)]];
                combine.Merge();
            }
            for (int i = leng + 4; i <= leng + 6; i++)
            {

                xlsSheet.Cells[4, 3 + i] = Sfreq_list[i - (leng + 4)] + "MHz";
            }

            
           
            //行列的合并
            //Range combine = xlsSheet.Rows["7:9",MisValue];
            //Range combine = xlsSheet.Columns.get_Range("G:",MisValue);
         
            //xlsbook.Save();
            app.DisplayAlerts = false;//不对是否取代原有文件进行询问
            xlsbook.SaveAs(System.Windows.Forms.Application.StartupPath+@"\tlp1.xlsx", MisValue, MisValue, MisValue, MisValue, MisValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, MisValue,MisValue, MisValue, MisValue,MisValue);
            xlsbook.Close();
            xlsbooks.Close();
            app.Quit();
            //释放掉多余的excel进程
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            app = null;
        }

        private void num_simpleButton_Click(object sender, EventArgs e)
        {
            this.Analysis_spreadsheetControl.Document.Worksheets[0].Cells["B" + 1].Value = this.Pnum_textEdit.Text;
        }
    }
}
