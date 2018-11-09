using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using com.suncreate.wb.agtlib;

namespace MeritNetwork
{
    public class NetWorkAnalyser
    {
        private AgtVisaCls agt = new AgtVisaCls();
        
        //string tmp = "";

        //IP地址连接
        public void NetWork_IP(string addr)
        {
            StringBuilder sb_addr = new StringBuilder();
            sb_addr.AppendFormat("TCPIP0::{0}::5025::SOCKET", addr);
            addr = sb_addr.ToString();
            agt.VisaOpen(addr);
        }
        //GPIB地址连接
        public void NetWork_GPIB(string addr)
        {
            StringBuilder sb_addr = new StringBuilder();
            sb_addr.AppendFormat("GPIB0::{0}::INSTR", addr);
            addr = sb_addr.ToString();
            agt.VisaOpen(addr);
        }

        public void NetWork_write(string addr)
        {
            agt.VisaWrite(addr);
        }

        public double NetWork_read(string addr)
        {
            return agt.VisaRead(addr);

        }

        public void NetWork_Close()
        {

            agt.VisaClose();
        }
        #region 常用方法定义

        //仪表初始化
        public void InitTest()
        {
            //定义测试名称
            agt.VisaWrite("SYST:PRES");
            agt.VisaWrite("DISPlay:WINDow1:STATE OFF");
            agt.VisaWrite("CALCulate:PARameter:DEFine 'MyMeas1',S11");
            agt.VisaWrite("CALCulate:PARameter:DEFine 'MyMeas2',S21");
            agt.VisaWrite("CALCulate:PARameter:DEFine 'MyMeas3',S21");
            agt.VisaWrite("CALCulate:PARameter:DEFine 'MyMeas4',S22");

            //打开窗口
            agt.VisaWrite("DISPlay:WINDow1:STATE ON");
            agt.VisaWrite("DISPlay:WINDow2:STATE ON");
            agt.VisaWrite("DISPlay:WINDow3:STATE ON");
            agt.VisaWrite("DISPlay:WINDow4:STATE ON");

            //显示对应测试曲线

            //S11 DB格式
            agt.VisaWrite("DISPlay:WINDow1:TRACe1:FEED 'MyMeas1'");
            agt.VisaWrite("CALCulate:PARameter:SELect 'MyMeas1'");
            agt.VisaWrite("CALCulate:Form MLOG");

            //S21 PHAS格式
            agt.VisaWrite("DISPlay:WINDow2:TRACe1:FEED 'MyMeas2'");
            agt.VisaWrite("CALCulate:PARameter:SELect 'MyMeas2'");
            agt.VisaWrite("CALCulate:Form PHAS");

            //S21 DB格式
            agt.VisaWrite("DISPlay:WINDow3:TRACe1:FEED 'MyMeas3'");
            agt.VisaWrite("CALCulate:PARameter:SELect 'MyMeas3'");
            agt.VisaWrite("CALCulate:Form MLOG");

            //S22 DB格式
            agt.VisaWrite("DISPlay:WINDow4:TRACe1:FEED 'MyMeas4'");
            agt.VisaWrite("CALCulate:PARameter:SELect 'MyMeas4'");
            agt.VisaWrite("CALCulate:Form MLOG");

            agt.VisaWrite("SENSe:SWE:MODE CONT");
          
            //CH1_S11_1,S11,CH1_S11_2,S21,CH1_S21_3,S21,CH1_S22_4,S22
        
        }

        //载入测试状态
        public void LoadState()
        {
            agt.VisaWrite(@"MMEM:Load  'd:\New Folder\DA_P_470_610.cst'");

        }

        public void LoadChuXiang()
        {
            agt.VisaWrite(@"MMEM:LOAD 'd:\1.1.csa'");
        }

        public void Makr1ON()
        {
            agt.VisaWrite("CALC:MARK1 ON");
        }
        //信息
        private string msg;
        public string Msg
        {
            get
            {
                msg = agt.VisaRead_STR("*IDN?");
                return msg;
            }
        }

        //测试初始化
        public void InitSetting(string startfreq, string stopfreq)
        {
            agt.VisaWrite("SENS:FREQ:STAR  " + startfreq);
            agt.VisaWrite("SENS:FREQ:STOP  " + stopfreq);
        }


        //读取数据
        private double[] trace_p;
        public double[] Trace_p
        {
            get
            {
                trace_p = agt.VisaReads("CALC1:DATA? FDATA");
                return trace_p;
            }
        }
        public void DataMemory()
        {
            agt.VisaWrite("CALC:MATH:MEM");
            agt.VisaWrite("CALC:MATH:FUNC DIV");
        }
        //选择测试窗口
        private int selectWin = 0;
        public int SelectWin
        {
            get { return selectWin; }
            set
            {
                if (value == 5)
	            {
                    selectWin = value;
                    agt.VisaWrite("CALCulate:PARameter:SELect 'CH1_S21_5'");
	            }
                else
                {
                    selectWin = value;
                    agt.VisaWrite("CALCulate:PARameter:SELect 'MyMeas" + selectWin + "'");
                }
            }
        }

        public void SelectChuXiang()
        {
            agt.VisaWrite("CALCulate:PARameter:SELect 'CH1_S11_1'");
        }

        //读取最大值
        private double maxAmp;
        public double MaxAmp
        {
            get
            {
                agt.VisaWrite("CALC:MARK1:MAX");
                System.Threading.Thread.Sleep(10);
                maxAmp = agt.VisaRead("CALC:MARK1:Y?");
                return maxAmp;
            }
            set { maxAmp = value; }
        }

        //读取最小值
        private double minAmp;
        public double MinAmp
        {
            get
            {
                agt.VisaWrite("CALC:MARK2:MIN");
                System.Threading.Thread.Sleep(10);
                minAmp = agt.VisaRead("CALC:MARK2:Y?");
                return minAmp;
            }
            set { minAmp = value; }
        }

        //读取X值对应Y值
        private double readAmp;
        public string ReadAmp
        {
            get
            {
                readAmp = agt.VisaRead("CALC:MARK1:Y?");
                string mark0 = readAmp.ToString("f2");
                return mark0;
            }
            set
            {
                agt.VisaWrite("CALC:MARK1:X " + value);
                System.Threading.Thread.Sleep(10);
            }
        }

        //private double mark1Y;
        public string Mark1Y
        {
            get
            {
                readAmp = agt.VisaRead("CALC:MARK1:Y?");
                string mark0 = readAmp.ToString("f3");
                return mark0;
            }
            set
            {
                agt.VisaWrite("CALC:MARK1:X " + value);
                System.Threading.Thread.Sleep(10);
            }
        }

        private double mark2Y;
        public string Mark2Y
        {
            get
            {
                mark2Y = agt.VisaRead("CALC:MARK2:Y?");
                mark2Y = mark2Y * 1000;
                string mark0 = mark2Y.ToString("f2");
                return mark0;
            }
            set
            {
                agt.VisaWrite("CALC:MARK2:X " + value);
                System.Threading.Thread.Sleep(10);
            }
        }

        //private double mark3Y;
        public double Mark3Y
        {
            get
            {

                maxAmp = agt.VisaRead("CALC:MARK3:Y?");
                return maxAmp;
            }
            set { maxAmp = value; }
        }


        #endregion


        //public double[] TestProject1()//端口驻波
        //{
         
        //    agt.VisaWrite("CALC:PAR:SEL 'CH1_S21_3'");
        //    agt.VisaWrite("CALC:FORM SWR");

        //    //double marker1 = Mark1Y;
        //    //double marker1Abs = System.Math.Abs(marker1);
        //    //AppendMsg("S21 SWR: " + marker1);

        //    agt.VisaWrite("CALC:PAR:SELect 'CH1_S22_4'");
        //    agt.VisaWrite("CALC:FORM SWR");//读取S22窗口

        //    double marker2 = Mark2Y;
        //    double marker2Abs = System.Math.Abs(marker2);
        //    //AppendMsg("S22 SWR: " + marker2);

        //    agt.VisaWrite("CALCulate:PARameter:SELect 'CH1_S21_3'");
        //    agt.VisaWrite("CALC:FORM MLOG");
        //    agt.VisaWrite("CALC:MARK3:X " + 9.5 + "Ghz");

        //    double marker3 = Mark3Y;
        //    double marker3Abs = System.Math.Abs(marker3);
        //    //AppendMsg("S21 MLOG: " + marker3);

        //    double[] list = new double[2];
        //    if (marker1Abs > marker2Abs)
        //    {
        //        if (marker1Abs > marker3Abs)
        //        {
        //            list[0] = marker1;
        //            list[1] = marker3;
        //            return list;
        //        }
        //        else
        //        {
        //            list[0] = marker3;
        //            list[1] = marker3;
        //            return list;
        //        }              
        //    }
        //    else
        //    {
        //        if (marker2Abs > marker3Abs)
        //        {
        //            list[0] = marker2;
        //            list[1] = marker3;
        //            return list;
        //        }
        //        else
        //        {
        //            list[0] = marker3;
        //            list[1] = marker3;
        //            return list;
        //        }   
        //    }
        //}



    }
}
