using CCWin;
using DevExpress.XtraSpreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MeritNetwork
{
    public partial class MessageForm : CCSkinMain
    {
        //Form1 frm1 = new Form1();
        double loss = 0;
        MYdelegate1 myde1;
        MYdelegate2 myde2;
        public MessageForm(MYdelegate1 md1,MYdelegate2 md2)
        {
            InitializeComponent();
            myde1 = md1;
            myde2 = md2;
        }

        /// <summary>
        /// 继续测试
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void simpleButton_GOON_Click(object sender, EventArgs e)
        {
            //this.DialogResult = DialogResult.OK;
        }

        private void simpleButton_Stop_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void skinTextBox_ID_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Work(myde1,myde2);
            }
        }

        private void Work(MYdelegate1 md1,MYdelegate2 md2)
        {
            int s1 = int.Parse(skinTextBox_ID.Text);
            int s2 = int.Parse(skinTextBox_ID.Text) + 1;
            if ((s1 % 2 == 1) ? true : false)
            {
                DialogResult r1 = MessageBox.Show("现在进行" + s1 + "," + s2 + "通道隔离度测试，请将矢网port2端口连接被测件端口" + s2 + ",按“是”进行测试，按否退出后进行线路连接", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                if ((int)r1 == 6)
                {
                    md1(s1);
                    loss = md2(Form1.fileM);
                }
                else
                {
                    this.DialogResult = DialogResult.Cancel;
                }
            }
            else
            {
                MessageBox.Show("请输入奇数！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
            }
        }

        private void sure_simpleButton_Click(object sender, EventArgs e)
        {
            Work(myde1,myde2);
        }
    }
}
