using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace com.suncreate.wb.agtlib
{

    public class AgtVisaCls
    {

        private int resourceManager = 0, viError;
        private int session = 0;

        public void VisaOpen(string addr)
        {
            viError = AgVisa32.viOpenDefaultRM(out resourceManager);
            if (viError != 0)
                throw new Exception("error:"+addr);

            viError = AgVisa32.viOpen(resourceManager, addr,
                AgVisa32.VI_NO_LOCK, AgVisa32.VI_TMO_IMMEDIATE, out session);
            if (viError != 0)
                throw new Exception("error:"+addr);
        }

        public void VisaClose()
        {
            AgVisa32.viClose(session);
            AgVisa32.viClose(resourceManager);
        }

        public void VisaWrite(string command)
        {

            viError = AgVisa32.viPrintf(session, command + "\n");
            if (viError != 0)
                throw new Exception("仪表连接错误！");

        }

        public double VisaRead(string command)
        {
            double dtmp;

            viError = AgVisa32.viPrintf(session, command + "\n");


            string res = "";
            AgVisa32.viRead(session, out res, 100);

            string[] resa = res.Split(',');
            res = resa[0];
            dtmp = double.Parse(res);

            return dtmp;

        }

        public string VisaRead_STR(string command)
        {
            viError = AgVisa32.viPrintf(session, command + "\n");
            string res = "";
            AgVisa32.viRead(session, out res, 100);
            return res;

        }

        public double[] VisaReads(string command)
        {
            double[] dtmp;

            viError = AgVisa32.viPrintf(session, command + "\n");


            string res = "";

            AgVisa32.viRead(session, out res, 1000000);
              
            string[] tmp = res.Split(',');
            dtmp = new double[tmp.Length];

            try
            {
                for (int i = 0; i < tmp.Length; i++)
                {
                    dtmp[i] = double.Parse(tmp[i]);
                }
            }
            catch
            {
                Console.WriteLine();
            }

            return dtmp;

        }







        public static bool MmemLoad(string addr, string filename)
        {
            try
            {

                int resourceManager = 0, viError;
                int session = 0;


                viError = AgVisa32.viOpenDefaultRM(out resourceManager);

                viError = AgVisa32.viOpen(resourceManager, addr.ToString(),
                    AgVisa32.VI_NO_LOCK, AgVisa32.VI_TMO_IMMEDIATE, out session);
                System.Threading.Thread.Sleep(100);
                viError = visa32.viPrintf(session, "MMEM:LOAD '" + filename + "'" + "\n");
                System.Threading.Thread.Sleep(100);
                AgVisa32.viClose(session);
                AgVisa32.viClose(resourceManager);
                return true;
            }
            catch
            {
                return false;
            }
        }


        public static void OpenPWR(string p, int dbm)
        {
            StringBuilder gpib = new StringBuilder();
            gpib = gpib.AppendFormat("TCPIP0::{0}::inst0::INSTR", p);

            int resourceManager = 0, viError;
            int session = 0;


            viError = AgVisa32.viOpenDefaultRM(out resourceManager);

            viError = AgVisa32.viOpen(resourceManager, gpib.ToString(),
                AgVisa32.VI_NO_LOCK, AgVisa32.VI_TMO_IMMEDIATE, out session);
            System.Threading.Thread.Sleep(100);
            viError = visa32.viPrintf(session, ":OUTPut ON" + "\n");
            viError = visa32.viPrintf(session, ":POWer " + dbm + " dBm" + "\n");
            System.Threading.Thread.Sleep(100);
            AgVisa32.viClose(session);
            AgVisa32.viClose(resourceManager);

        }

        public static void ClsPWR(string p)
        {
            StringBuilder gpib = new StringBuilder();
            gpib = gpib.AppendFormat("TCPIP0::{0}::inst0::INSTR", p);

            int resourceManager = 0, viError;
            int session = 0;


            viError = AgVisa32.viOpenDefaultRM(out resourceManager);

            viError = AgVisa32.viOpen(resourceManager, gpib.ToString(),
                AgVisa32.VI_NO_LOCK, AgVisa32.VI_TMO_IMMEDIATE, out session);
            System.Threading.Thread.Sleep(100);
            viError = AgVisa32.viPrintf(session, ":OUTPut OFF" + "\n");
            System.Threading.Thread.Sleep(100);
            AgVisa32.viClose(session);
            AgVisa32.viClose(resourceManager);
        }

        public static double ReadDB(string addr, string window, string m1)
        {

            int resourceManager = 0, viError;
            int session = 0;


            viError = AgVisa32.viOpenDefaultRM(out resourceManager);

            viError = AgVisa32.viOpen(resourceManager, addr.ToString(),
                AgVisa32.VI_NO_LOCK, AgVisa32.VI_TMO_IMMEDIATE, out session);
            System.Threading.Thread.Sleep(100);

            viError = AgVisa32.viPrintf(session, "CALC:PAR:SEL '" + window + "'" + "\n");

            viError = AgVisa32.viPrintf(session, "CALC:" + m1 + ":Y?" + "\n");
            string res = "";
            AgVisa32.viRead(session, out res, 100);

            System.Threading.Thread.Sleep(100);
            AgVisa32.viClose(session);
            AgVisa32.viClose(resourceManager);
            double dtmp = double.Parse(res.ToString().Split(',')[0]);
            return dtmp;

        }


        public static void SetMark(string p, string m1, double value)
        {
            StringBuilder gpib = new StringBuilder();
            gpib = gpib.AppendFormat("TCPIP0::{0}::inst0::INSTR", p);

            int resourceManager = 0, viError;
            int session = 0;


            viError = AgVisa32.viOpenDefaultRM(out resourceManager);

            viError = AgVisa32.viOpen(resourceManager, gpib.ToString(),
                AgVisa32.VI_NO_LOCK, AgVisa32.VI_TMO_IMMEDIATE, out session);
            System.Threading.Thread.Sleep(100);
            viError = visa32.viPrintf(session, "CALC:" + m1 + ":X " + value + "MHz" + "\n");
            System.Threading.Thread.Sleep(100);
            AgVisa32.viClose(session);
            AgVisa32.viClose(resourceManager);

        }

        public static void VisaWrite(string ip, string command)
        {
            StringBuilder gpib = new StringBuilder();
            gpib = gpib.AppendFormat("TCPIP0::{0}::inst0::INSTR", ip);

            int resourceManager = 0, viError;
            int session = 0;


            viError = AgVisa32.viOpenDefaultRM(out resourceManager);

            viError = AgVisa32.viOpen(resourceManager, gpib.ToString(),
                AgVisa32.VI_NO_LOCK, AgVisa32.VI_TMO_IMMEDIATE, out session);
            System.Threading.Thread.Sleep(100);
            viError = visa32.viPrintf(session, command + "\n");
            System.Threading.Thread.Sleep(100);
            AgVisa32.viClose(session);
            AgVisa32.viClose(resourceManager);
        }
       

    }
}
