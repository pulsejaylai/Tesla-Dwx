using System;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Text;
//using Microsoft.Office.Core;
using Excel=Microsoft.Office.Interop.Excel;
using Ivi.Visa.Interop;
using System.Threading;










namespace DWXTEST
{
    public partial class Form1 : Form
    {
       // SerialPort sp = new SerialPort("COM3");
        SerialPort sp;
        FileStream fs;
        FileStream savedata;
        FileStream sndata;
        StreamWriter sw2;
        StreamWriter sw3;
        StreamReader sr2;
        StreamWriter checksn;
        Thread thread1, thread2,thread3;
        UInt32 rvalue1817_12, rvalue1817_34, rvalue1818, rvalue1877;
        double fard;
        delegate void mydelegate();
        //   string DutAddr = "USB0::0x0471::0x2883::QF40900001::0::INSTR";
        string DutAddr = "ASRL1::INSTR";
        string DutAddr2 = "ASRL3::INSTR";
        Ivi.Visa.Interop.ResourceManager rm = new Ivi.Visa.Interop.ResourceManager(); //Open up a new resource manager
        Ivi.Visa.Interop.FormattedIO488 myDmm = new Ivi.Visa.Interop.FormattedIO488();
        string configtxt,savepath,testresult,SN,sncheck,model,path2,days,equ,equgpib2883,gpconfig,ticonfig,tistr,relayconfig,IOtype,M1877Lx;
        string[] itemstr;
      //  string[] gp2883;
       
        int testdo,x,t17,dwxjudge,testitemno,itemprocess,dwxi,th2883i,th2839i,thjudge;
        bool comtrue,teststatue,filenew;
        Int16 hcard = -1;
        [DllImport(@"C:\Windows\System32\LPCI7250.dll", EntryPoint = "Init_Card7296", CallingConvention = CallingConvention.Cdecl)]
        public static extern Int16 Init7296();
        [DllImport(@"C:\Windows\System32\LPCI7250.dll", EntryPoint = "Card7296P1AOUTPUT", CallingConvention = CallingConvention.Cdecl)]
        public static extern void card7296configP1A(Int16 cardno);
        [DllImport(@"C:\Windows\System32\LPCI7250.dll", EntryPoint = "Card7296P1ARELAYOUTPUT", CallingConvention = CallingConvention.Cdecl)]
        public static extern void card7296P1Arelayout(Int16 cardno, UInt32 data);
        [DllImport(@"C:\Windows\System32\LPCI7250.dll", EntryPoint = "Card7296Realease", CallingConvention = CallingConvention.Cdecl)]
        public static extern void cardrealease(Int16 cardno);
        [DllImport("kernel32.dll")]
        public static extern uint GetTickCount();
        public static void Delay(uint ms)
        {
            uint start = GetTickCount();
            while (GetTickCount() - start < ms)
            {
                System.Windows.Forms.Application.DoEvents();
            }
        }  
        
        
        
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            System.Diagnostics.Process[] myProcesses = System.Diagnostics.Process.GetProcessesByName("DWXTEST");//获取指定的进程名   
            if (myProcesses.Length > 1) //如果可以获取到知道的进程名则说明已经启动
            {
                MessageBox.Show("程序已启动！");
                System.Windows.Forms.Application.Exit();              //关闭系统
           
            }

            checkBox1.Checked = true;
            equ = "";
            equgpib2883 = "";
            dwxjudge = 0;
            thjudge = 0;
            t17 = 99;
            string comstr;
            string[] p;
            int pi, i,ii;
             p=new string[6];
             itemstr = new string[6];
            pi = 0;
             x = 0;
            comtrue = true;
             teststatue = false;
             this.label1.Font = new System.Drawing.Font("隶书", 24, FontStyle.Bold); //第一个是字体，第二个大小，第三个是样式，
             this.label1.ForeColor = Color.Blue;// 颜色 
             this.snshow.Font = new System.Drawing.Font("隶书", 24, FontStyle.Bold); //第一个是字体，第二个大小，第三个是样式，
             this.snshow.ForeColor = Color.Blue;// 颜色 
             snshow.Text = "SN";
            //  this.label1.Text = "Ready"; 
            
            
            //   sp = new SerialPort("com3");
            // s=SerialPort.GetPortNames();
          //  MessageBox.Show(s[0]);
            configtxt =  "d:\\dwconfig.txt";
            savepath = "d:\\DWTEST";
            sncheck = "d:\\DWTEST\\sncheck.txt";
            gpconfig = "d:\\gpconfig.txt";
            ticonfig = "d:\\Itemconfig.txt";
            relayconfig = "d:\\relayconfig.txt";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            
          
           
            
            if (!File.Exists(sncheck))
            {
                sndata = new FileStream(sncheck, FileMode.Create);
                checksn = new StreamWriter(sndata);
                checksn.Flush();
                checksn.Close();
                //  filenew = true;
            }

            if (!File.Exists(relayconfig))
            {
                fs = new FileStream(relayconfig, FileMode.Create);

                sw2 = new StreamWriter(fs);
                sw2.Write("PCI7296" + "\r\n");

                sw2.Flush();
                //关闭流
                sw2.Close();
                fs.Close();
                IOtype = "PCI7296";
                hcard = Init7296();
                if (hcard < 0)
                {
                    MessageBox.Show("IOCard Init Fail="+hcard.ToString());
                    TEST.Enabled = false;
                }
                else
                {
                    card7296configP1A(hcard);
                    card7296P1Arelayout(hcard,0X00);
                }
           
            
            }
            else
            {
                fs = new FileStream(relayconfig, FileMode.Open);
                sr2 = new StreamReader(fs);
                while (!sr2.EndOfStream)
                {
                    IOtype = sr2.ReadLine();

                    //  MessageBox.Show(comstr);
                }

                //  MessageBox.Show(comstr);
                fs.Close();
                if (IOtype.IndexOf("7296") != -1)
                {
                    hcard = Init7296();
                    if (hcard < 0)
                    {
                        MessageBox.Show("I/O Card Init Fail=" + hcard.ToString());
                    }
                    else
                    {
                        card7296configP1A(hcard);
                        card7296P1Arelayout(hcard, 0X00);
                    }


                }
           
          
          }
            
            
            
            if (!File.Exists(gpconfig))
            {
                fs = new FileStream(gpconfig, FileMode.Create);

                sw2 = new StreamWriter(fs);
                sw2.Write("ASRL3::INSTR" + "\r\n");

                sw2.Flush();
                //关闭流
                sw2.Close();
                fs.Close();

                equgpib2883 = "ASRL3::INSTR";
                //    myDmm.IO = (IMessage)rm.Open(DutAddr, AccessMode.NO_LOCK, 2000, ""); //Open up a handle to the DMM with a 2 second timeout
                //  myDmm.IO.Timeout = 3000;

            }
            else
            {
                fs = new FileStream(gpconfig, FileMode.Open);
                sr2 = new StreamReader(fs);
                while (!sr2.EndOfStream)
                {
                    equgpib2883 = sr2.ReadLine();

                    //  MessageBox.Show(comstr);
                }

                //  MessageBox.Show(comstr);
                fs.Close();
            }
            
            
            
            
            /*  else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
            }*/
            
            
            
           /* path2 = DateTime.Now.ToString("yyyy");
            savepath = savepath + "\\" +model+"\\"+ path2 + ".txt";            
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write( "\r\n"+"\r\n"+"\r\n");
            }*/
            if (!File.Exists(ticonfig))
            {
                fs = new FileStream(ticonfig, FileMode.Create);

                sw2 = new StreamWriter(fs);
                sw2.Write("2" + "\r\n");
                sw2.Write("DWX_1817_12" + "\r\n");
                sw2.Write("TH2883_1817_34" + "\r\n");
                sw2.Flush();
                //关闭流
                sw2.Close();
                fs.Close();
                sr2 = new StreamReader(fs);
                while (!sr2.EndOfStream)
                {
                    comstr = sr2.ReadLine();
                    p[pi] = comstr;
                    pi = pi + 1;
                    //  MessageBox.Show(comstr);
                }

                //  MessageBox.Show(comstr);
                fs.Close();
                //    MessageBox.Show(p[0]);
                testitemno = int.Parse(p[0]);
             for(ii=0;ii<testitemno;ii++)
             { itemstr[ii] = p[ii + 1]; }
                // sp = new SerialPort(p[0]);
             //   DutAddr = "ASRL" + p[0].Substring(3, 1) + "::INSTR";
                //    myDmm.IO = (IMessage)rm.Open(DutAddr, AccessMode.NO_LOCK, 2000, ""); //Open up a handle to the DMM with a 2 second timeout
                //  myDmm.IO.Timeout = 3000;

            }
            else
            {
                fs = new FileStream(ticonfig, FileMode.Open);
                sr2 = new StreamReader(fs);
                while (!sr2.EndOfStream)
                {
                    comstr = sr2.ReadLine();
                    p[pi] = comstr;
                    pi = pi + 1;
                    //  MessageBox.Show(comstr);
                }

                //  MessageBox.Show(comstr);
                fs.Close();
              //      MessageBox.Show(p[0]);
                testitemno = int.Parse(p[0]);
                for (ii = 0; ii < testitemno; ii++)
                { itemstr[ii] = p[ii + 1]; }

                //  MessageBox.Show(DutAddr);
                //  sp = new SerialPort(p[0]);     
            }










            pi = 0;
            if (!File.Exists(configtxt))
            {
                fs = new FileStream(configtxt, FileMode.Create);
            
                sw2 = new StreamWriter(fs);
                sw2.Write("COM3"+"\r\n");
                sw2.Write("9600"+"\r\n");
                sw2.Write("8" + "\r\n");
                sw2.Write("None" + "\r\n");
                sw2.Write("1" + "\r\n");
                sw2.Write("None" + "\r\n");
                sw2.Flush();
                //关闭流
                sw2.Close();
                fs.Close();
                sr2 = new StreamReader(fs);
                while (!sr2.EndOfStream)
                {
                    comstr = sr2.ReadLine();
                    p[pi] = comstr;
                    pi = pi + 1;
                    //  MessageBox.Show(comstr);
                }

                //  MessageBox.Show(comstr);
                fs.Close();
                //    MessageBox.Show(p[0]);
                sp = new SerialPort(p[0]);
                DutAddr = "ASRL" + p[0].Substring(3, 1) + "::INSTR";
            //    myDmm.IO = (IMessage)rm.Open(DutAddr, AccessMode.NO_LOCK, 2000, ""); //Open up a handle to the DMM with a 2 second timeout
              //  myDmm.IO.Timeout = 3000;
            
            }
            else
            {
                fs = new FileStream(configtxt, FileMode.Open);
                sr2 = new StreamReader(fs);
                while (!sr2.EndOfStream)
                {
                    comstr = sr2.ReadLine();
                    p[pi] = comstr;
                    pi = pi + 1;
                    //  MessageBox.Show(comstr);
                }
                
              //  MessageBox.Show(comstr);
                fs.Close();
            //    MessageBox.Show(p[0]);
                sp = new SerialPort(p[0]);
              //  MessageBox.Show(p[0].Substring(3, 1));
               // DutAddr = "ASRL1::INSTR";
                DutAddr = "ASRL" + p[0].Substring(3, 1) + "::INSTR";
               
                //  MessageBox.Show(DutAddr);
                    //  sp = new SerialPort(p[0]);     
            }
            
            
            try
            {
               
                sp.Open();
               
                //sp.Close();
               // MessageBox.Show("OK");
            }
            catch (InvalidOperationException eee)
            {
              //  MessageBox.Show("yidakai");
            }

            catch (Exception ee)
            {
                MessageBox.Show(ee.Message);
                comtrue = false;
                TEST.Enabled = false;
                this.label1.ForeColor = Color.Red;
                this.label1.Text = "Com Err"; 
            }
            sp.Close();
          
            
            
            if (comtrue == true)
            {
                this.label1.Text = "Ready"; 
                try
                {
                    sp.BaudRate = int.Parse(p[1]);
                    sp.DataBits = int.Parse(p[2]);
                    if (p[3] == "even")
                    {
                        sp.Parity = Parity.Even;
                    }
                    if (p[3] == "odd")
                    {
                        sp.Parity = Parity.Odd;
                    }
                    if (p[3] == "mark")
                    {
                        sp.Parity = Parity.Mark;
                    }
                    if (p[3] == "none")
                    {
                        sp.Parity = Parity.None;
                    }
                    if (p[3] == "space")
                    {
                        sp.Parity = Parity.Space;
                    }
                    if (p[4] == "1")
                    {
                        sp.StopBits = StopBits.One;
                    }
                    if (p[4] == "2")
                    {
                        sp.StopBits = StopBits.Two;
                    }
                    if (p[5] == "none")
                    {
                        sp.Handshake = Handshake.None;
                    }
                    if (p[5] == "send")
                    {
                        sp.Handshake = Handshake.RequestToSend;
                    }
                    if (p[5] == "xonxoff")
                    {
                        sp.Handshake = Handshake.XOnXOff;
                    }
                }
                catch (Exception ee)
                {
                    MessageBox.Show(ee.Message);
                    comtrue = false;
                    TEST.Enabled = false;
                }

           /*     sp.Open();
                if (sp.IsOpen)
                {
                    //  t17 = 17;
                    sp.NewLine = "\r\n";
                    //   sp.WriteLine("B,1817    .034");
                    sp.WriteLine("U");

                }
                sp.Close();*/
            
            
            }

           

            sp.ReceivedBytesThreshold = 1;
            sp.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(sp_DataReceived);
          
          
            
            dataGridView1.ColumnCount = 12;
            dataGridView1.Columns[0].HeaderText = "SN";
            dataGridView1.Columns[0].Width = 255;
            dataGridView1.Columns[1].HeaderText = "Area(Freq)";
            dataGridView1.Columns[1].Width = 95;
            dataGridView1.Columns[2].HeaderText = "Lapic(Cp)";
            dataGridView1.Columns[2].Width = 95;
            dataGridView1.Columns[3].HeaderText = "Result";
            dataGridView1.Columns[3].Width = 95;
            dataGridView1.Columns[4].HeaderText = "Turn";
            dataGridView1.Columns[4].Width = 95;
            dataGridView1.Columns[5].HeaderText = "Lx";
            dataGridView1.Columns[5].Width = 95;
            dataGridView1.Columns[6].HeaderText = "Lk";
            dataGridView1.Columns[6].Width = 95;
            dataGridView1.Columns[7].HeaderText = "Cx";
            dataGridView1.Columns[7].Width = 95;
            dataGridView1.Columns[7].HeaderText = "Q";
            dataGridView1.Columns[7].Width = 95;
            dataGridView1.Columns[8].HeaderText = "DCR";
            dataGridView1.Columns[8].Width = 95;
            dataGridView1.Columns[9].HeaderText = "ACR";
            dataGridView1.Columns[9].Width = 95;
            dataGridView1.Columns[10].HeaderText = "Zx";
            dataGridView1.Columns[10].Width = 95;
            dataGridView1.Columns[11].HeaderText = "POL";
            dataGridView1.Columns[11].Width = 95;
            
            
            
            dataGridView1.ReadOnly = true;
         //   this.ControlBox = false;
            this.dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            string year, month, day;
            int iyear, imonth, iday,yy,xx;
            year = DateTime.Now.ToString("yyyy");
          //  MessageBox.Show(year);
            iyear = int.Parse(year);
            if ((iyear % 4 == 0 && iyear % 100 != 0) || (iyear % 400 == 0))
            {
                yy = 1;
            }
            else
            {
                yy = 0;
            }
            month = DateTime.Now.ToString("MM");
            //  MessageBox.Show(year);
            imonth = int.Parse(month);
            xx = 0;
            switch (imonth)
            {
                case 12:
                    xx = 30 + 31 + 30 + 31 + 31 + 30 + 31 + 28 + yy + 31 + 30 + 31;
                    break;
                case 11:
                    xx = 31 + 30 + 31 + 31 + 30 + 31 + 28 + yy + 31 + 30 + 31;
                    break;
                case 10:
                    xx = 30 + 31 + 31 + 30 + 31 + 28 + yy + 31 + 30 + 31;                
                    break;
                case 9:
                    xx = 31 + 31 + 30 + 31 + 28 + yy + 31 + 30 + 31;
                    break;
                case 8:
                    xx = 31 + 30 + 31 + 28 + yy + 31 + 30 + 31;
                    break;
                case 7:
                    xx = 30 + 31 + 28 + yy + 31 + 30 + 31;
                    break;
                case 6:
                    xx = 31 + 28 + yy + 31 + 30+31;
                    break;
                case 5:
                    xx = 31 + 28 + yy + 31+30;
                    break;
                case 4: 
                    xx = 31 +28+yy+31;
                    break;
                case 3: 
                    xx = 28 + yy+31;
                    break;
                case 2:
                    xx =  31;
                    break;
                case 1:
                   // xx = 0;
                    xx = 0; 
                    break;
                default:
                    xx = 0;
                    break;



            }
            day = DateTime.Now.ToString("dd");
            iday = int.Parse(day);
            days = (xx + iday).ToString();
           // MessageBox.Show(days);
            if (days.Length == 1)
            { days = "00" + days; }
            if (days.Length == 2)
            { days = "0" + days; }
            textBox1.Text = days;
        
        
        }

        private void comSetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ComConfig ff = new ComConfig();
            ff.TransfEvent += frm_TransfEvent;
            DialogResult ddr = ff.ShowDialog();
        }

      //  private void TEST_Change(object sender, EventArgs e)
        void frm_TransfEvent()
        {
            this.TEST.Enabled = true;
        }

        void Sn_TransfEvent(string value)
        {
            SN = value;
        }

        void Statue_TransintfEvent(int valuet)
        {
            testdo = valuet;
        }

        void Meas_Deg()
        {
            string result, aread, lppd;
            int xi;
            string[] seq;
            string[] seq2;
            string[] seq3;
            int ind, mi, ibuf, i11;
            // double fard;
            result = "";
            myDmm.WriteString("FETC?", true);
            result = myDmm.ReadString();
            //  MessageBox.Show(result);
            xi = 0;
            seq3 = new string[3];
            seq = result.Split(',');
            foreach (string azu in seq)
            {
                seq3[xi] = azu;
                xi++;

            }
            ind = seq3[1].IndexOf("E");
            aread = seq3[1].Substring(0, ind);
            lppd = seq3[1].Substring(ind + 1, seq3[1].Length - ind - 1);
            //lppd = lppd.Substring(lppd.Length-1,1);
            //  MessageBox.Show(aread);
            // MessageBox.Show(lppd);
            if (lppd.Substring(0, 1) == "-")
            {
                fard = double.Parse(aread);
                mi = int.Parse(lppd.Substring(lppd.Length - 1, 1));
                fard = fard / System.Math.Pow(10, mi) * 100;
                //   deg = fard;
            }
            if (lppd.Substring(0, 1) == "+")
            {
                fard = double.Parse(aread);
                mi = int.Parse(lppd.Substring(lppd.Length - 1, 1));
                fard = fard * System.Math.Pow(10, mi) * 100;
                //  deg = fard;
            }

            // return deg;
        }

        void Thead_Test3()
        {
            //  do
            // {

            Threadx3();
            //}while(fx==1);

        }
       
        void Threadx3()
        {
            string result;
            FileStream fs;
            if (dataGridView1.InvokeRequired == false)
            {
            myDmm.WriteString("*IDN?\n", true);
            result = myDmm.ReadString();
            if (result.IndexOf("T") == -1)
            {
                this.label1.Text = "GPIB Err";
            }
            if (result.IndexOf("T") != -1)
            {
                fs = new FileStream("d:\\TH2829RESULT.bin", FileMode.CreateNew);
                fs.Close();
              /*  myDmm.WriteString("TRIG\n", true);
myDmm.ReadList(*/

           
            
            
            }        
        
}
            else
            {
                // MessageBox.Show("Mid2");
                mydelegate mytest3 = new mydelegate(Threadx3);


                dataGridView1.BeginInvoke(mytest3);

            }        
      

}
        
        void Thead_Test2()
        {
            //  do
            // {

            Threadx2();
            //}while(fx==1);

        }


        void Threadx2()
        {
            string result, aread, lppd, common;
            result = "pass";
            aread = "";
            lppd = "";
            string[] seq;
            string[] seq2;
            string[] seq3;
            string[] fhz;
            int xi, ind, mi, ibuf, i11, i12, i13, i14;
            int[] ifhz;

            int rhz;
            fard = 0;
            mi = 0;
            FileStream fs = null;
            Encoding encoder = Encoding.UTF8;
            byte[] bytes;
            if (dataGridView1.InvokeRequired == false)
            {


                if (model == "M1818")
                {
                    /*  fhz = new string[620];
                      ifhz = new int[620];
                      for (ibuf = 0; ibuf < 620; ibuf++)
                      {
                          ifhz[ibuf] = ibuf + 380;
                          fhz[ibuf] = ifhz[ibuf].ToString();
                      }*/
                    //  myDmm.WriteString("TRIG", true);
                    //  Delay(500);
                    //    myDmm.WriteString("FETC:CRES?", true);
                    ibuf = 0;
                    myDmm.WriteString("FREQ 400KHZ", true);
                    myDmm.WriteString("TRIG", true);
                    myDmm.WriteString("TRIG", true);
                    Meas_Deg();
                    if (fard <= 0)
                    {
                        //  MessageBox.Show(fard.ToString());
                        testresult = "ng";
                        result = "fail";
                        this.dataGridView1.Rows.Add();
                        this.dataGridView1.Rows[x].Cells[0].Value = SN;
                        this.dataGridView1.Rows[x].Cells[1].Value = fard.ToString();
                        this.dataGridView1.Rows[x].Cells[3].Value = "FAIL";
                        this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.Red;
                        dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                        this.dataGridView1.Rows[x].Selected = true;
                        this.label1.ForeColor = Color.Red;
                        this.label1.Text = "FAIL";
                        th2839i = 1;
                        this.TEST.Enabled = true;
                    }
                    myDmm.WriteString("FREQ 1MHZ", true);
                    myDmm.WriteString("TRIG", true);
                    myDmm.WriteString("TRIG", true);
                    Meas_Deg();
                    if (fard > 0)
                    {
                        testresult = "ng";
                        result = "fail";
                        this.dataGridView1.Rows.Add();
                        this.dataGridView1.Rows[x].Cells[0].Value = SN;
                        this.dataGridView1.Rows[x].Cells[1].Value = fard.ToString();
                        this.dataGridView1.Rows[x].Cells[3].Value = "FAIL";
                        this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.Red;
                        dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                        this.dataGridView1.Rows[x].Selected = true;
                        this.label1.ForeColor = Color.Red;
                      
                        this.label1.Text = "FAIL";
                        th2839i = 1;
                        this.TEST.Enabled = true;
                    }
                    if (result == "pass")
                    {
                        testresult = "ok";
                        myDmm.WriteString("FREQ 400KHZ", true);
                        myDmm.WriteString("TRIG", true);
                        myDmm.WriteString("TRIG", true);
                        Meas_Deg();
                        i11 = 400;
                        do
                        {

                            common = "FREQ " + i11.ToString() + "KHZ";
                            myDmm.WriteString(common, true);
                            myDmm.WriteString("TRIG", true);
                            myDmm.WriteString("TRIG", true);
                            Meas_Deg();
                            i11 = i11 + 100;
                            Delay(3);
                        } while (fard >= 0);
                        /* if(fard==0)
                         {

                         }*/
                        i12 = i11 - 200;
                        do
                        {

                            common = "FREQ " + i12.ToString() + "KHZ";
                            myDmm.WriteString(common, true);
                            myDmm.WriteString("TRIG", true);
                            myDmm.WriteString("TRIG", true);
                            Meas_Deg();
                            i12 = i12 + 50;
                            Delay(3);
                        } while (fard >= 0);

                        i13 = i12 - 100;
                        do
                        {

                            common = "FREQ " + i13.ToString() + "KHZ";
                            myDmm.WriteString(common, true);
                            myDmm.WriteString("TRIG", true);
                            myDmm.WriteString("TRIG", true);
                            Meas_Deg();
                            i13 = i13 + 10;
                            Delay(3);
                        } while (fard >= 0);

                        i14 = i13 - 20;
                        do
                        {

                            common = "FREQ " + i14.ToString() + "KHZ";
                            myDmm.WriteString(common, true);
                            myDmm.WriteString("TRIG", true);
                            myDmm.WriteString("TRIG", true);
                            Meas_Deg();
                            i14 = i14 + 1;
                            Delay(3);
                        } while (fard >= 0);

                        i14 = i14 - 2;
                        this.dataGridView1.Rows.Add();
                        this.dataGridView1.Rows[x].Cells[0].Value = SN;
                        this.dataGridView1.Rows[x].Cells[1].Value = i14.ToString() + "KHz";
                        this.dataGridView1.Rows[x].Cells[3].Value = "PASS";
                        this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.White;
                        dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                        this.dataGridView1.Rows[x].Selected = true;
                        savedata = new FileStream(savepath, FileMode.Append);
                        sw3 = new StreamWriter(savedata);
                        if (model == "M181734")
                        { sw3.Write("SN: " + SN + "-34" + "   "); }
                        else
                        { sw3.Write("SN: " + SN + "   "); }
                        sw3.Write("Hz: " + i14 + "KHz" + "   ");
                        //sw3.Write("AREA:" + seq2[0].Substring(seq2[0].Length - seq2[0].IndexOf("=")) + "\r\n");
                        // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("=") - 1);
                        // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("="));
                        // ps = ps.Replace(",", "");
                        //   sw3.Write("LAPLC:" + seq2[3] + "\r\n");
                        sw3.Write("RESULT: " + "PASS" + "\r\n" );
                        sw3.Flush();
                        //关闭流
                        sw3.Close();
                        savedata.Close();
                        if (itemprocess == testitemno - 1)
                        {
                            fs = File.OpenWrite("d:\\2839TEST\\sncheck.txt");
                            fs.Position = fs.Length;
                            bytes = encoder.GetBytes(SN + "\r\n");
                            fs.Write(bytes, 0, bytes.Length);
                            fs.Close();
                            this.label1.ForeColor = Color.Green;
                            this.label1.Text = "PASS";
                            this.TEST.Enabled = true;
                            if (checkBox1.Checked == true)
                            {
                                th2839i = 1;
                                TEST.PerformClick();
                            }
                        }

                        if (itemprocess != testitemno - 1)
                        { th2839i = 1; }
                  
    
    
    }
                    //   Delay(200);
                    //  myDmm.WriteString("FETC?", true);
                    /*   do
                       {
                      
                           common = "FREQ " + fhz[ibuf] + "KHZ";
                           myDmm.WriteString(common, true);
                           myDmm.WriteString("TRIG", true);
                           myDmm.WriteString("TRIG", true);
                           //   Delay(100);
                           myDmm.WriteString("FETC?", true);
                           result = myDmm.ReadString();
                           //  MessageBox.Show(result);
                           xi = 0;
                           seq3 = new string[3];
                           seq = result.Split(',');
                           foreach (string azu in seq)
                           {
                               seq3[xi] = azu;
                               xi++;

                           }
                           ind = seq3[1].IndexOf("E");
                           aread = seq3[1].Substring(0, ind);
                           lppd = seq3[1].Substring(ind + 1, seq3[1].Length - ind - 1);
                           //lppd = lppd.Substring(lppd.Length-1,1);
                           if (lppd.Substring(0, 1) == "-")
                           {
                               fard = double.Parse(aread);
                               mi = int.Parse(lppd.Substring(lppd.Length - 1, 1));
                               fard = fard / System.Math.Pow(10, mi) * 100;
                           }
                           if (lppd.Substring(0, 1) == "+")
                           {
                               fard = double.Parse(aread);
                               mi = int.Parse(lppd.Substring(lppd.Length - 1, 1));
                               fard = fard * System.Math.Pow(10, mi) * 100;
                           }

                      
                        
                        
                           ibuf++;
                           Delay(5);
                       } while ((fard>=0)&&(ibuf<620));*/
                    //   myDmm.WriteString("FETC:CRES?", true);
                    // result = myDmm.ReadString();
                    //MessageBox.Show(result);
                    /*  }*/


                    /*  xi = 0;
                      seq2 = new string[7];
                      seq = result.Split(',');
                      foreach (string azu in seq)
                      {
                          seq2[xi] = azu;
                          xi++;

                      }
                      ind = seq2[1].IndexOf("E");
                      aread = seq2[1].Substring(0, ind);
                      lppd = seq2[1].Substring(ind + 1, seq2[1].Length - ind - 1);
                      //lppd = lppd.Substring(lppd.Length-1,1);
                      if (lppd.Substring(0, 1) == "-")
                      {
                          fard = double.Parse(aread);
                          mi = int.Parse(lppd.Substring(lppd.Length - 1, 1));
                          fard = fard / System.Math.Pow(10, mi) * 100;
                      }
                      if (lppd.Substring(0, 1) == "+")
                      {
                          fard = double.Parse(aread);
                          mi = int.Parse(lppd.Substring(lppd.Length - 1, 1));
                          fard = fard * System.Math.Pow(10, mi) * 100;
                      }*/
                    //  MessageBox.Show(fard.ToString());
                    /*   this.dataGridView1.Rows.Add();
                       this.dataGridView1.Rows[x].Cells[0].Value = SN;
                       this.dataGridView1.Rows[x].Cells[1].Value = fhz[ibuf];*/
                    // this.dataGridView1.Rows[x].Cells[2].Value = seq2[3];

                    /*   if ((ifhz[ibuf]>=400)&&(ibuf<220))
                       {
                           this.dataGridView1.Rows[x].Cells[3].Value = "PASS";
                           this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.White;
                           dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                           this.dataGridView1.Rows[x].Selected = true;
                           savedata = new FileStream(savepath, FileMode.Append);
                           sw3 = new StreamWriter(savedata);
                           if (model == "M181734")
                           { sw3.Write("SN:" + SN + "-34" + "\r\n"); }
                           else
                           { sw3.Write("SN:" + SN + "\r\n"); }
                           sw3.Write("Hz:" + fhz[ibuf] + "\r\n");
                           //sw3.Write("AREA:" + seq2[0].Substring(seq2[0].Length - seq2[0].IndexOf("=")) + "\r\n");
                           // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("=") - 1);
                           // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("="));
                           // ps = ps.Replace(",", "");
                        //   sw3.Write("LAPLC:" + seq2[3] + "\r\n");
                           sw3.Write("RESULT:" + "PASS" + "\r\n" + "\r\n" + "\r\n");
                           sw3.Flush();
                           //关闭流
                           sw3.Close();
                           savedata.Close();
                           fs = File.OpenWrite("d:\\2839TEST\\sncheck.txt");
                           fs.Position = fs.Length;
                           bytes = encoder.GetBytes(SN + "\r\n");
                           fs.Write(bytes, 0, bytes.Length);
                           fs.Close();
                           this.label1.ForeColor = Color.Green;
                           this.label1.Text = "PASS";
                           this.TEST.Enabled = true;
                           TEST.PerformClick();


                           // ssresult = "PASS";
                       }*/
                    /* if (ifhz[ibuf] <400)
                     {
                         this.dataGridView1.Rows[x].Cells[3].Value = "FAIL";
                         this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.Red;
                         dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                         this.dataGridView1.Rows[x].Selected = true;
                         this.label1.ForeColor = Color.Red;
                         this.label1.Text = "FAIL";
                         this.TEST.Enabled = true;
                         //  ssresult = "FAIL";
                     }*/
                    x++;
                }

                if (model == "M1829")
                {
                    int yy;
                    myDmm.WriteString("FREQ 10MHZ", true);
                    for (yy = 0; yy < 3; yy++)
                    {
                        myDmm.WriteString("TRIG", true);
                        myDmm.WriteString("TRIG", true);
                        Delay(100);
                        //    myDmm.WriteString("FETC:CRES?", true);
                        /*  do
                          {
                              result = myDmm.ReadString();
                              //  MessageBox.Show(result);
                          } while (result.IndexOf("END") == -1);*/
                        myDmm.WriteString("FETC?", true);
                        result = myDmm.ReadString();
                        //MessageBox.Show(result);
                        xi = 0;
                        seq2 = new string[7];
                        seq = result.Split(',');
                        foreach (string azu in seq)
                        {
                            seq2[xi] = azu;
                            xi++;

                        }
                        ind = seq2[0].IndexOf("E");
                        aread = seq2[0].Substring(0, ind);
                        lppd = seq2[0].Substring(ind + 1, seq2[1].Length - ind - 1);
                        //lppd = lppd.Substring(lppd.Length-1,1);
                        //   MessageBox.Show(lppd.Substring(lppd.Length - 1, 1));
                        if (lppd.Substring(0, 1) == "-")
                        {
                            fard = double.Parse(aread);
                            mi = int.Parse(lppd.Substring(lppd.Length - 2, 2));
                            fard = fard / System.Math.Pow(10, mi) * System.Math.Pow(10, 12);
                        }
                        if (lppd.Substring(0, 1) == "+")
                        {
                            fard = double.Parse(aread);
                            mi = int.Parse(lppd.Substring(lppd.Length - 2, 2));
                            fard = fard * System.Math.Pow(10, mi) * System.Math.Pow(10, 12);
                        }
                        //  MessageBox.Show(fard.ToString());
                        this.dataGridView1.Rows.Add();
                        this.dataGridView1.Rows[x].Cells[0].Value = SN;
                        //  this.dataGridView1.Rows[x].Cells[1].Value = fard.ToString() + "%";
                        this.dataGridView1.Rows[x].Cells[2].Value = fard.ToString();

                        if ((fard < 33) && (fard > 0))
                        {
                            testresult = "ok";
                            this.dataGridView1.Rows[x].Cells[3].Value = "PASS";
                            this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.White;
                            dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                            this.dataGridView1.Rows[x].Selected = true;
                            savedata = new FileStream(savepath, FileMode.Append);
                            sw3 = new StreamWriter(savedata);
                            if (yy == 0)
                            { sw3.Write("SN: " + SN + "-12" + "   "); }
                            if (yy == 1)
                            { sw3.Write("SN: " + SN + "-34" + "   "); }
                            if (yy == 2)
                            { sw3.Write("SN: " + SN + "-56" + "   "); }
                            sw3.Write("Cp: " + fard.ToString() + "   ");
                            //sw3.Write("AREA:" + seq2[0].Substring(seq2[0].Length - seq2[0].IndexOf("=")) + "\r\n");
                            // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("=") - 1);
                            // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("="));
                            // ps = ps.Replace(",", "");
                            //  sw3.Write("LAPLC:" + seq2[3] + "\r\n");
                            sw3.Write("RESULT: " + "PASS" + "\r\n");
                            sw3.Flush();
                            //关闭流
                            sw3.Close();
                            savedata.Close();
                            if (yy == 0)
                            { MessageBox.Show("Test 3-4"); }
                            if (yy == 1)
                            { MessageBox.Show("Test 5-6"); }
                            if (yy == 2)
                            {
                                if (itemprocess == testitemno - 1)
                                {
                                    fs = File.OpenWrite("d:\\2839TEST\\sncheck.txt");
                                    fs.Position = fs.Length;
                                    bytes = encoder.GetBytes(SN + "\r\n");
                                    fs.Write(bytes, 0, bytes.Length);
                                    fs.Close();
                                    this.label1.ForeColor = Color.Green;
                                    this.label1.Text = "PASS";
                                    th2839i = 1;
                                    this.TEST.Enabled = true;
                                    if (checkBox1.Checked == true)
                                    { TEST.PerformClick(); }
                                }
                                if (itemprocess != testitemno - 1)
                                { th2839i = 1; }  
                             
                                /*   this.TEST.Enabled = true;
                                   TEST.PerformClick();*/
                            }

                            // ssresult = "PASS";
                        }
                        if ((fard >= 33) || (fard < 0))
                        {
                            testresult = "ng";
                            this.dataGridView1.Rows[x].Cells[3].Value = "FAIL";
                            this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                            this.dataGridView1.Rows[x].Selected = true;
                            this.label1.ForeColor = Color.Red;
                            this.label1.Text = "FAIL";
                            this.TEST.Enabled = true;
                            th2839i = 1;
                            break;
                            //  ssresult = "FAIL";
                        }

                        x++;
                    }

                   
                }


                if (model == "M0075")
                {
                    int yy;
                    myDmm.WriteString("FREQ 10MHZ", true);
                    for (yy = 0; yy < 2; yy++)
                    {
                        myDmm.WriteString("TRIG", true);
                        myDmm.WriteString("TRIG", true);
                        Delay(100);
                        //    myDmm.WriteString("FETC:CRES?", true);
                        /*  do
                          {
                              result = myDmm.ReadString();
                              //  MessageBox.Show(result);
                          } while (result.IndexOf("END") == -1);*/
                        myDmm.WriteString("FETC?", true);
                        result = myDmm.ReadString();
                        //MessageBox.Show(result);
                        xi = 0;
                        seq2 = new string[7];
                        seq = result.Split(',');
                        foreach (string azu in seq)
                        {
                            seq2[xi] = azu;
                            xi++;

                        }
                        ind = seq2[0].IndexOf("E");
                        aread = seq2[0].Substring(0, ind);
                        lppd = seq2[0].Substring(ind + 1, seq2[1].Length - ind - 1);
                        //lppd = lppd.Substring(lppd.Length-1,1);
                        //   MessageBox.Show(lppd.Substring(lppd.Length - 1, 1));
                        if (lppd.Substring(0, 1) == "-")
                        {
                            fard = double.Parse(aread);
                            mi = int.Parse(lppd.Substring(lppd.Length - 2, 2));
                            fard = fard / System.Math.Pow(10, mi) * System.Math.Pow(10, 12);
                        }
                        if (lppd.Substring(0, 1) == "+")
                        {
                            fard = double.Parse(aread);
                            mi = int.Parse(lppd.Substring(lppd.Length - 2, 2));
                            fard = fard * System.Math.Pow(10, mi) * System.Math.Pow(10, 12);
                        }
                        //  MessageBox.Show(fard.ToString());
                        this.dataGridView1.Rows.Add();
                        this.dataGridView1.Rows[x].Cells[0].Value = SN;
                        //  this.dataGridView1.Rows[x].Cells[1].Value = fard.ToString() + "%";
                        this.dataGridView1.Rows[x].Cells[2].Value = fard.ToString();

                        if ((fard < 33) && (fard > 0))
                        {
                            testresult = "ok";
                               
                            this.dataGridView1.Rows[x].Cells[3].Value = "PASS";
                            this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.White;
                            dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                            this.dataGridView1.Rows[x].Selected = true;
                            savedata = new FileStream(savepath, FileMode.Append);
                            sw3 = new StreamWriter(savedata);
                            if (yy == 0)
                            { sw3.Write("SN: " + SN + "-12" + "   "); }
                            if (yy == 1)
                            { sw3.Write("SN: " + SN + "-34" + "   "); }
                            /*  if (yy == 2)
                              { sw3.Write("SN:" + SN + "-56" + "\r\n"); }*/
                            sw3.Write("Cp: " + fard.ToString() + "\r\n");
                            //sw3.Write("AREA:" + seq2[0].Substring(seq2[0].Length - seq2[0].IndexOf("=")) + "\r\n");
                            // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("=") - 1);
                            // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("="));
                            // ps = ps.Replace(",", "");
                            //  sw3.Write("LAPLC:" + seq2[3] + "\r\n");
                            sw3.Write("RESULT: " + "PASS" + "\r\n");
                            sw3.Flush();
                            //关闭流
                            sw3.Close();
                            savedata.Close();
                            if (yy == 0)
                            { MessageBox.Show("Test 3-4"); }
                            /*     if (yy == 1)
                                 { MessageBox.Show("Test 5-6"); }*/
                            if (yy == 1)
                            {
                                if (itemprocess == testitemno - 1)
                                {
                                    fs = File.OpenWrite("d:\\2839TEST\\sncheck.txt");
                                    fs.Position = fs.Length;
                                    bytes = encoder.GetBytes(SN + "\r\n");
                                    fs.Write(bytes, 0, bytes.Length);
                                    fs.Close();
                                    this.label1.ForeColor = Color.Green;
                                    this.label1.Text = "PASS";
                                    th2839i = 1;
                                    this.TEST.Enabled = true;
                                    if (checkBox1.Checked == true)
                                    { TEST.PerformClick(); }
                                }
                                    /*   this.TEST.Enabled = true;
                                   TEST.PerformClick();*/
                                if (itemprocess != testitemno - 1)
                                { th2839i = 1; }
                           
                           }

                            // ssresult = "PASS";
                        }
                        if ((fard >= 33) || (fard < 0))
                        {
                            this.dataGridView1.Rows[x].Cells[3].Value = "FAIL";
                            this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                            this.dataGridView1.Rows[x].Selected = true;
                            this.label1.ForeColor = Color.Red;
                            this.label1.Text = "FAIL";
                            this.TEST.Enabled = true;
                            th2839i = 1;
                            break;
                            //  ssresult = "FAIL";
                        }

                        x++;
                    }

                  
                }




























            }


            else
            {
                // MessageBox.Show("Mid2");
                mydelegate mytest2 = new mydelegate(Threadx2);


                dataGridView1.BeginInvoke(mytest2);

            }








        }     
      
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        void Thead_Test()
        {
            //  do
            // {

            Threadx();
            //}while(fx==1);

        }
        void Threadx()
        {
            string result, aread, lppd,sn2;
            result = "";
            aread = "";
            lppd = "";
            string[] seq;
            string[] seq2;
            int xi, ind, mi;
            double fard;
            fard = 0;
            mi = 0;
            sn2 = "";
            FileStream fs = null;
            Encoding encoder = Encoding.UTF8;
            byte[] bytes;
            if (dataGridView1.InvokeRequired == false)
            {


                if ((model == "M1818") || (model == "M1817") || (model == "M1816") || (model == "M181734") || (model == "M181634") || (model == "M1877") || (model == "M187734") || (model == "M1877SHORT"))
                {
                    if (model.IndexOf("M1817") !=-1)
                    {
                        if (itemstr[itemprocess].IndexOf("_12") != -1)
                        {
                           // myDmm.WriteString("MMEM:LOAD:STAT 4", true);
                        /*    myDmm.WriteString("IVOLT:VOLT 3800V", true);
                            myDmm.WriteString("IVOLT:VADJ ON", true);
                            myDmm.WriteString("IVOLT:TIMP 5", true);
                           
                            myDmm.WriteString("IVOLT:EIMP 0", true);
                           
                            myDmm.WriteString("COMP:AREA ON", true);
                            
                            myDmm.WriteString("COMP:AREA:RANG 0,2196", true);
                            
                            myDmm.WriteString("COMP:AREA:DIFF 10", true);
                            
                            myDmm.WriteString("COMP:DIFF OFF", true);
                            
                            myDmm.WriteString("COMP:CORONA ON", true);
                            
                            myDmm.WriteString("COMP:CORONA:RANG 0,2196", true);
                            
                            myDmm.WriteString("COMP:CORONA:DIFF 250", true);
                            
                            myDmm.WriteString("COMP:PHAS OFF", true);*/
                            
                            myDmm.WriteString("MMEM:LOAD:STAT 4", true);
                            card7296P1Arelayout(hcard, rvalue1817_12);
                        }
                        if (itemstr[itemprocess].IndexOf("_34") != -1)
                        {
                          //  myDmm.WriteString("MMEM:LOAD:STAT 3", true);
                          /*  myDmm.WriteString("IVOLT:VOLT 2000V", true);
                            myDmm.WriteString("IVOLT:VADJ ON", true);
                            myDmm.WriteString("IVOLT:TIMP 5", true);                         
                            myDmm.WriteString("IVOLT:EIMP 0", true);                          
                            myDmm.WriteString("COMP:AREA ON", true);          
                            myDmm.WriteString("COMP:AREA:RANG 0,894", true);               
                            myDmm.WriteString("COMP:AREA:DIFF 10", true);               
                            myDmm.WriteString("COMP:DIFF OFF", true);          
                            myDmm.WriteString("COMP:CORONA ON", true);            
                            myDmm.WriteString("COMP:CORONA:RANG 0,894", true);           
                            myDmm.WriteString("COMP:CORONA:DIFF 250", true);        
                            myDmm.WriteString("COMP:PHAS OFF", true);    */ 
                            myDmm.WriteString("MMEM:LOAD:STAT 3", true);
                            card7296P1Arelayout(hcard, rvalue1817_34);
                        }
                    //    Delay(2000);
                        myDmm.WriteString("TRIG", true);
                    }
                    if (model.IndexOf("M1818") != -1)
                    {
                        myDmm.WriteString("MMEM:LOAD:STAT 1", true);
                        card7296P1Arelayout(hcard, rvalue1818);
                        myDmm.WriteString("TRIG", true); 
                    }
                   
                    if (model.IndexOf("1816") != -1)
                    {
                     //   MessageBox.Show(itemstr[itemprocess]);
                        if (itemstr[itemprocess].IndexOf("_Open") != -1)
                        {
                            if (itemstr[itemprocess].IndexOf("4.5") != -1)
                            { myDmm.WriteString("MMEM:LOAD:STAT 8", true); }
                            if (itemstr[itemprocess].IndexOf("3.8") != -1)
                            { myDmm.WriteString("MMEM:LOAD:STAT 6", true); }
                            if (itemstr[itemprocess].IndexOf("4.9") != -1)
                            { myDmm.WriteString("MMEM:LOAD:STAT 10", true); }
                            if (itemstr[itemprocess].IndexOf("5.3") != -1)
                            { myDmm.WriteString("MMEM:LOAD:STAT 14", true); }
                            if (itemstr[itemprocess].IndexOf("5.9") != -1)
                            { myDmm.WriteString("MMEM:LOAD:STAT 12", true); }
                            
                            
                            
                            
                            card7296P1Arelayout(hcard, rvalue1877); }
if (itemstr[itemprocess].IndexOf("_Short") != -1)
{
    MessageBox.Show("短路");
    if (itemstr[itemprocess].IndexOf("4.5") != -1)
    { myDmm.WriteString("MMEM:LOAD:STAT 9", true); }
    if (itemstr[itemprocess].IndexOf("3.8") != -1)
    { myDmm.WriteString("MMEM:LOAD:STAT 7", true); }
    if (itemstr[itemprocess].IndexOf("4.9") != -1)
    { myDmm.WriteString("MMEM:LOAD:STAT 11", true); }
    if (itemstr[itemprocess].IndexOf("5.3") != -1)
    { myDmm.WriteString("MMEM:LOAD:STAT 15", true); }
    if (itemstr[itemprocess].IndexOf("5.9") != -1)
    { myDmm.WriteString("MMEM:LOAD:STAT 13", true); }
    
    
    
    card7296P1Arelayout(hcard, rvalue1877);
}
myDmm.WriteString("TRIG", true);
                    }
                    
                    if (model.IndexOf("M1877") != -1)
                    {
                        if (itemstr[itemprocess].IndexOf("_Open") != -1)
                        {

                            if (M1877Lx == "3.8")
                            { myDmm.WriteString("MMEM:LOAD:STAT 6", true); }
                            if (M1877Lx == "4.5")
                            { myDmm.WriteString("MMEM:LOAD:STAT 8", true); }
                            if (M1877Lx == "4.9")
                            { myDmm.WriteString("MMEM:LOAD:STAT 10", true); }
                            if (M1877Lx == "5.3")
                            { myDmm.WriteString("MMEM:LOAD:STAT 14", true); }
                            if (M1877Lx == "5.9")
                            { myDmm.WriteString("MMEM:LOAD:STAT 12", true); }
                            card7296P1Arelayout(hcard, rvalue1877);
                        }
                        if (itemstr[itemprocess].IndexOf("_Short") != -1)
                        {

                            MessageBox.Show("短路");
                            if (M1877Lx == "3.8")
                            { myDmm.WriteString("MMEM:LOAD:STAT 7", true); }
                            if (M1877Lx == "4.5")
                            { myDmm.WriteString("MMEM:LOAD:STAT 9", true); }
                            if (M1877Lx == "4.9")
                            { myDmm.WriteString("MMEM:LOAD:STAT 11", true); }
                            if (M1877Lx == "5.3")
                            { myDmm.WriteString("MMEM:LOAD:STAT 15", true); }
                            if (M1877Lx == "5.9")
                            { myDmm.WriteString("MMEM:LOAD:STAT 13", true); }
                            card7296P1Arelayout(hcard, rvalue1877);

                        }
                        //    Delay(2000);
                        myDmm.WriteString("TRIG", true);
                    }    
                
                    
                    
                    Delay(1500);
                    //    myDmm.WriteString("FETC:CRES?", true);
                  /*  do
                    {
                        
                        result = myDmm.ReadString();
                        //  MessageBox.Show(result);
                    } while (result.IndexOf("END") == -1);*/
                    myDmm.WriteString("FETC:CRES?", true);
                    result = myDmm.ReadString();
                    //MessageBox.Show(result);
                    xi = 0;
                    seq2 = new string[7];
                    seq = result.Split(',');
                    foreach (string azu in seq)
                    {
                        seq2[xi] = azu;
                        xi++;

                    }
                    ind = seq2[1].IndexOf("E");
                    aread = seq2[1].Substring(0, ind);
                    lppd = seq2[1].Substring(ind + 1, seq2[1].Length - ind - 1);
                    //lppd = lppd.Substring(lppd.Length-1,1);
                    if (lppd.Substring(0, 1) == "-")
                    {
                        fard = double.Parse(aread);
                        mi = int.Parse(lppd.Substring(lppd.Length - 1, 1));
                        fard = fard / System.Math.Pow(10, mi) * 100;
                    }
                    if (lppd.Substring(0, 1) == "+")
                    {
                        fard = double.Parse(aread);
                        mi = int.Parse(lppd.Substring(lppd.Length - 1, 1));
                        fard = fard * System.Math.Pow(10, mi) * 100;
                    }
                    //  MessageBox.Show(fard.ToString());
                    this.dataGridView1.Rows.Add();
                    this.dataGridView1.Rows[x].Cells[0].Value = SN;
                    this.dataGridView1.Rows[x].Cells[1].Value = fard.ToString() + "%";
                    this.dataGridView1.Rows[x].Cells[2].Value = seq2[3];

                    if (seq2[0] == "1")
                    {
                        testresult = "ok";
                        this.dataGridView1.Rows[x].Cells[3].Value = "PASS";
                        this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.White;
                        dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                        this.dataGridView1.Rows[x].Selected = true;
                        savedata = new FileStream(savepath, FileMode.Append);
                        sw3 = new StreamWriter(savedata);
                        if ((itemstr[itemprocess].IndexOf("34") != -1) || (itemstr[itemprocess].IndexOf("Short") != -1))
                        {
                            sn2 = SN;
                            sw3.Write("SN: " + sn2 + "-34" + "   ");
                        }
                        else
                        {
                          //  if (itemprocess == 0)
                          //  {
                                sn2 = SN;
                                sw3.Write("SN: " + sn2 + "   ");
                                snshow.Text =  sn2;
                          //  }
                         /*   if (itemprocess != 0)
                            {
                                SnInput snbox = new SnInput();
                                snbox.TransfEvent += Sn_TransfEvent;
                                snbox.TransintfEvent += Statue_TransintfEvent;
                                snbox.ModelSet = model;
                                snbox.DaySet = textBox1.Text.Substring(0, 3);
                                snbox.EquSet = equ;
                                if (checkBox1.Checked == true)
                                { DialogResult ddr = snbox.ShowDialog(); }
                                sn2 = SN;
                                sw3.Write("SN: " + sn2 + "   ");
                                snshow.Text = sn2;
                                this.dataGridView1.Rows[x].Cells[0].Value = sn2;
                            }
                        */
                        
                        
                        
                        
                        }
                        sw3.Write("AREA: " + fard.ToString() + "   ");
                        //sw3.Write("AREA:" + seq2[0].Substring(seq2[0].Length - seq2[0].IndexOf("=")) + "\r\n");
                        // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("=") - 1);
                        // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("="));
                        // ps = ps.Replace(",", "");
                        sw3.Write("LAPLC: " + seq2[3] + "   ");
                        sw3.Write("RESULT: " + "PASS" + "\r\n");
                        sw3.Flush();
                        //关闭流
                        sw3.Close();
                        savedata.Close();
                        if (itemprocess == testitemno - 1)
                        {
                            card7296P1Arelayout(hcard, 0X00);
                            fs = File.OpenWrite("d:\\DWTEST\\sncheck.txt");
                            fs.Position = fs.Length;
                            bytes = encoder.GetBytes(sn2 + "\r\n");
                            fs.Write(bytes, 0, bytes.Length);
                            fs.Close();
                          
                            this.label1.ForeColor = Color.Green;
                            this.label1.Text = "PASS";
                            this.TEST.Enabled = true;
                            if (checkBox1.Checked == true)
                            {
                                th2883i = 1;
                                TEST.PerformClick();
                            }
                        }
                        if (itemprocess != testitemno - 1)
                        { th2883i = 1; }
                       
                        // ssresult = "PASS";
                    }
                    if (seq2[0] == "0")
                    {
                        card7296P1Arelayout(hcard, 0X00);
                        testresult = "ng";
                        this.dataGridView1.Rows[x].Cells[3].Value = "FAIL";
                        this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.Red;
                        dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                        this.dataGridView1.Rows[x].Selected = true;
                        this.label1.ForeColor = Color.Red;
                        this.label1.Text = "FAIL";
                        this.TEST.Enabled = true;
                        //  ssresult = "FAIL";
                        if (itemstr[itemprocess].IndexOf("34") != -1)
                        { sn2 = SN; }
                        else
                        {
                            if (itemprocess == 0)
                            {
                                sn2 = SN;                        
                                snshow.Text = sn2;
                            }
                        /*    if (itemprocess != 0)
                            {
                                SnInput snbox = new SnInput();
                                snbox.TransfEvent += Sn_TransfEvent;
                                snbox.TransintfEvent += Statue_TransintfEvent;
                                snbox.ModelSet = model;
                                snbox.DaySet = textBox1.Text.Substring(0, 3);
                                snbox.EquSet = equ;
                                if (checkBox1.Checked == true)
                                { DialogResult ddr = snbox.ShowDialog(); }
                                sn2 = SN;                          
                                snshow.Text = sn2;
                                this.dataGridView1.Rows[x].Cells[0].Value = sn2;
                            }
                        */
                        }
                        th2883i = 1;
                    }
                    x++;
                 //   th2883i = 1;
                }

                if (model == "M1829")
                {
                    int yy;
                    myDmm.WriteString("MMEM:LOAD:STAT 2", true);
                    card7296P1Arelayout(hcard, 0x0C);
                    Delay(200);
                    for (yy = 0; yy < 3; yy++)
                    {
                      
                        myDmm.WriteString("TRIG", true);
                        Delay(200);
                        //    myDmm.WriteString("FETC:CRES?", true);
                      /*  do
                        {
                            result = myDmm.ReadString();
                            //  MessageBox.Show(result);
                        } while (result.IndexOf("END") == -1);*/
                        myDmm.WriteString("FETC:CRES?", true);
                        result = myDmm.ReadString();
                        //MessageBox.Show(result);
                        xi = 0;
                        seq2 = new string[7];
                        seq = result.Split(',');
                        foreach (string azu in seq)
                        {
                            seq2[xi] = azu;
                            xi++;

                        }
                        ind = seq2[1].IndexOf("E");
                        aread = seq2[1].Substring(0, ind);
                        lppd = seq2[1].Substring(ind + 1, seq2[1].Length - ind - 1);
                        //lppd = lppd.Substring(lppd.Length-1,1);
                        if (lppd.Substring(0, 1) == "-")
                        {
                            fard = double.Parse(aread);
                            mi = int.Parse(lppd.Substring(lppd.Length - 1, 1));
                            fard = fard / System.Math.Pow(10, mi) * 100;
                        }
                        if (lppd.Substring(0, 1) == "+")
                        {
                            fard = double.Parse(aread);
                            mi = int.Parse(lppd.Substring(lppd.Length - 1, 1));
                            fard = fard * System.Math.Pow(10, mi) * 100;
                        }
                        //  MessageBox.Show(fard.ToString());
                        this.dataGridView1.Rows.Add();
                        this.dataGridView1.Rows[x].Cells[0].Value = SN;
                        this.dataGridView1.Rows[x].Cells[1].Value = fard.ToString() + "%";
                        this.dataGridView1.Rows[x].Cells[2].Value = seq2[3];

                        if (seq2[0] == "1")
                        {
                            testresult = "ok";
                            this.dataGridView1.Rows[x].Cells[3].Value = "PASS";
                            this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.White;
                            dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                            this.dataGridView1.Rows[x].Selected = true;
                            savedata = new FileStream(savepath, FileMode.Append);
                            sw3 = new StreamWriter(savedata);
                            if (yy == 0)
                            { sw3.Write("SN: " + SN + "-12" + "   "); }
                            if (yy == 1)
                            { sw3.Write("SN: " + SN + "-34" + "   "); }
                            if (yy == 2)
                            { sw3.Write("SN: " + SN + "-56" + "   "); }
                            sw3.Write("AREA: " + fard.ToString() + "   ");
                            //sw3.Write("AREA:" + seq2[0].Substring(seq2[0].Length - seq2[0].IndexOf("=")) + "\r\n");
                            // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("=") - 1);
                            // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("="));
                            // ps = ps.Replace(",", "");
                            sw3.Write("LAPLC: " + seq2[3] + "   ");
                            sw3.Write("RESULT: " + "PASS" + "\r\n");
                            sw3.Flush();
                            //关闭流
                            sw3.Close();
                            savedata.Close();
                            if (yy == 0)
                            { MessageBox.Show("Test 3-4"); }
                            if (yy == 1)
                            { MessageBox.Show("Test 5-6"); }
                            if (yy == 2)
                            {
                                if (itemprocess == testitemno - 1)
                                {
                                   // card7296P1Arelayout(hcard, 0X00);
                                    fs = File.OpenWrite("d:\\DWTEST\\sncheck.txt");
                                    fs.Position = fs.Length;
                                    bytes = encoder.GetBytes(SN + "\r\n");
                                    fs.Write(bytes, 0, bytes.Length);
                                    fs.Close();
                                    this.label1.ForeColor = Color.Green;
                                    this.label1.Text = "PASS";
                                    this.TEST.Enabled = true;
                                    if (checkBox1.Checked == true)
                                    {
                                        th2883i = 1;
                                        TEST.PerformClick(); 
                                    }
                                    /*  this.TEST.Enabled = true;
                                      TEST.PerformClick();*/
                                    
                                }
                                if (itemprocess != testitemno - 1)
                                { th2883i = 1; }
                            }

                            // ssresult = "PASS";
                        }
                        if (seq2[0] == "0")
                        {
                            card7296P1Arelayout(hcard, 0X00);
                            testresult = "ng";
                            this.dataGridView1.Rows[x].Cells[3].Value = "FAIL";
                            this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                            this.dataGridView1.Rows[x].Selected = true;
                            this.label1.ForeColor = Color.Red;
                            this.label1.Text = "FAIL";
                            this.TEST.Enabled = true;
                            th2883i = 1;
                            break;
                            //  ssresult = "FAIL";
                        }

                        x++;
                    }

                  
                }


                if (model == "M0075")
                {
                    int yy;

                    for (yy = 0; yy < 2; yy++)
                    {
                        myDmm.WriteString("TRIG", true);
                        Delay(200);
                        //    myDmm.WriteString("FETC:CRES?", true);
                        do
                        {
                            result = myDmm.ReadString();
                            //  MessageBox.Show(result);
                        } while (result.IndexOf("END") == -1);
                        myDmm.WriteString("FETC:CRES?", true);
                        result = myDmm.ReadString();                      
                        //MessageBox.Show(result);
                        xi = 0;
                        seq2 = new string[7];
                        seq = result.Split(',');
                        foreach (string azu in seq)
                        {
                            seq2[xi] = azu;
                            xi++;

                        }
                        ind = seq2[1].IndexOf("E");
                        aread = seq2[1].Substring(0, ind);
                        lppd = seq2[1].Substring(ind + 1, seq2[1].Length - ind - 1);
                        //lppd = lppd.Substring(lppd.Length-1,1);
                        if (lppd.Substring(0, 1) == "-")
                        {
                            fard = double.Parse(aread);
                            mi = int.Parse(lppd.Substring(lppd.Length - 1, 1));
                            fard = fard / System.Math.Pow(10, mi) * 100;
                        }
                        if (lppd.Substring(0, 1) == "+")
                        {
                            fard = double.Parse(aread);
                            mi = int.Parse(lppd.Substring(lppd.Length - 1, 1));
                            fard = fard * System.Math.Pow(10, mi) * 100;
                        }
                        //  MessageBox.Show(fard.ToString());
                        this.dataGridView1.Rows.Add();
                        this.dataGridView1.Rows[x].Cells[0].Value = SN;
                        this.dataGridView1.Rows[x].Cells[1].Value = fard.ToString() + "%";
                        this.dataGridView1.Rows[x].Cells[2].Value = seq2[3];

                        if (seq2[0] == "1")
                        {
                            testresult = "ok";
                            this.dataGridView1.Rows[x].Cells[3].Value = "PASS";
                            this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.White;
                            dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                            this.dataGridView1.Rows[x].Selected = true;
                            savedata = new FileStream(savepath, FileMode.Append);
                            sw3 = new StreamWriter(savedata);
                            if (yy == 0)
                            { sw3.Write("SN: " + SN + "-12" + "   "); }
                            if (yy == 1)
                            { sw3.Write("SN: " + SN + "-34" + "   "); }
                            /*   if (yy == 2)
                               { sw3.Write("SN:" + SN + "-56" + "\r\n"); }*/
                            sw3.Write("AREA: " + fard.ToString() + "   ");
                            //sw3.Write("AREA:" + seq2[0].Substring(seq2[0].Length - seq2[0].IndexOf("=")) + "\r\n");
                            // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("=") - 1);
                            // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("="));
                            // ps = ps.Replace(",", "");
                            sw3.Write("LAPLC: " + seq2[3] + "   ");
                            sw3.Write("RESULT: " + "PASS" + "\r\n" );
                            sw3.Flush();
                            //关闭流
                            sw3.Close();
                            savedata.Close();
                            if (yy == 0)
                            { MessageBox.Show("Test 3-4"); }
                            /*    if (yy == 1)
                                { MessageBox.Show("Test 5-6"); }*/
                            if (yy == 1)
                            {

                                if (itemprocess == testitemno - 1)
                                {
                                    card7296P1Arelayout(hcard, 0X00);
                                    fs = File.OpenWrite("d:\\DWTEST\\sncheck.txt");
                                fs.Position = fs.Length;
                                bytes = encoder.GetBytes(SN + "\r\n");
                                fs.Write(bytes, 0, bytes.Length);
                                fs.Close();
                                this.label1.ForeColor = Color.Green;
                                this.label1.Text = "PASS";
                                this.TEST.Enabled = true;
                                if (checkBox1.Checked == true)
                                {
                                    th2883i = 1;
                                    TEST.PerformClick();
                                }
                                
                                
                                }
                                /*  this.TEST.Enabled = true;
                                  TEST.PerformClick();*/
                                if (itemprocess != testitemno - 1)
                                { th2883i = 1; }
                            }
                        
                            // ssresult = "PASS";
                        }
                        if (seq2[0] == "0")
                        {
                            card7296P1Arelayout(hcard, 0X00);
                            testresult = "ng";
                            this.dataGridView1.Rows[x].Cells[3].Value = "FAIL";
                            this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.Red;
                            dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                            this.dataGridView1.Rows[x].Selected = true;
                            this.label1.ForeColor = Color.Red;
                            this.label1.Text = "FAIL";
                            this.TEST.Enabled = true;
                            th2883i = 1;
                            break;
                            //  ssresult = "FAIL";
                        }

                        x++;
                    }

               
             
                }


















            }


            else
            {
                // MessageBox.Show("Mid2");
                mydelegate mytest = new mydelegate(Threadx);


                dataGridView1.BeginInvoke(mytest);

            }








        }
        
        
        
        private void TEST_Click(object sender, EventArgs e)
        {
            TEST.Enabled = false;
            t17 = 0;
            dwxjudge = 0;
            testresult = "ok";
            dwxi = 0; th2839i = 0; th2883i = 0;
            string result, result2;
            int xi;
            SnInput snbox = new SnInput();
            snbox.TransfEvent += Sn_TransfEvent;
            snbox.TransintfEvent += Statue_TransintfEvent;
            snbox.ModelSet = model;
            snbox.DaySet = textBox1.Text.Substring(0,3);
            snbox.EquSet = equ;
            if (checkBox1.Checked == true)
            { DialogResult ddr = snbox.ShowDialog(); }
            if (checkBox1.Checked == false)
            { testdo = 1; }
                if (testdo == 1)
            {
               //SN=
                if (checkBox1.Checked == true)
                { snshow.Text = SN; }
                if (checkBox1.Checked == false)
                { 
                    snshow.Text = "NO SN";
                SN = "0000";
                }
                    this.label1.Text = "TEST";
                this.label1.ForeColor = Color.Yellow;
                result = "";
                teststatue = true;
                for (xi = 0; xi < testitemno; xi++)
                {
                   /* if (testitemno == 2)
                    { dwxi = 0; th2839i = 0; th2883i = 0; }*/
                  //  if (testitemno == 1)
                   // { dwxi = 1; th2839i = 1; th2883i = 1; }
                    dwxi = 0; th2839i = 0; th2883i = 0;
                    itemprocess = xi;
                    if (testresult != "ok")
                    { break; }
                    if (itemstr[xi].IndexOf("DWX")!=-1)
                    {
                        dwxjudge = 1;
                        try
                        {
                            sp.Open();
                            if (itemstr[xi].IndexOf("1876") == -1)
                            {
                                if (sp.IsOpen)
                                {
                                    if ((itemstr[xi].IndexOf("1829") != -1) || (itemstr[xi].IndexOf("0075") != -1))
                                    { MessageBox.Show("测试1，2脚"); }
                                    
                                    sp.NewLine = "\r\n";
                                    sp.WriteLine("S");


                                }
                            }
                            if (itemstr[xi].IndexOf("1876") != -1)
                            {
                                if (sp.IsOpen)
                                {
                                   /* if (xi == 0)
                                    {
                                        sp.NewLine = "\r\n";
                                        sp.WriteLine("S");
                                    }*/
                                    if (itemstr[xi].IndexOf("Short") != -1)
                                    {
                                        MessageBox.Show("短路");
                                        sp.NewLine = "\r\n";
                                        sp.WriteLine("S,1816    .034");
                                    }
                                    if (itemstr[xi].IndexOf("Open") != -1)
                                    {
                                      //  MessageBox.Show("开路");
                                        sp.NewLine = "\r\n";
                                        sp.WriteLine("S,1816    .012");
                                    }
                                
                                
                                
                                
                                }
                            }
                        
                    
                        
                        
                        
                        
                        
                        }
                        catch (Exception ee)
                        {
                            MessageBox.Show(ee.Message);

                        }
                      //  MessageBox.Show("begin");

                        if (testitemno != 1)
                        {
                            while (dwxi == 0)
                            { Delay(5); }
                        }
                            //    MessageBox.Show("end");
                    
                    
                    }

                    if (itemstr[xi].IndexOf("TH2883") != -1)
                    {
                        dwxjudge = 0;
                        if ((itemstr[xi].IndexOf("1829") != -1) || (itemstr[xi].IndexOf("0075") != -1))
                        { MessageBox.Show("测试1，2脚"); }
                   
                     
                        thread1 = new Thread(new ThreadStart(Thead_Test));
                        thread1.IsBackground = true;
                        thread1.Start();
                        if (testitemno != 1)
                        {
                            while (th2883i == 0)
                            { Delay(5); }
                        }



                 
                    
                    
                    
                    }
                    if (itemstr[xi].IndexOf("TH2839") != -1)
                    {
                        dwxjudge = 0;
                        if ((itemstr[xi].IndexOf("1829") != -1) || (itemstr[xi].IndexOf("0075") != -1))
                        { MessageBox.Show("测试1，2脚"); }
                        thread2 = new Thread(new ThreadStart(Thead_Test2));
                        thread2.IsBackground = true;
                        thread2.Start();
                        if (testitemno != 1)
                        {
                            while (th2839i == 0)
                            { Delay(5); }
                        }



                    }

                }
            
          
            
          
            
            
            
            }
          
            else
            {
                teststatue = false;
                TEST.Enabled = true;
                this.label1.Text = "Cancel";
                this.label1.ForeColor = Color.Blue;
            }

          //  sp.Close();
        }

        private void sp_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            string result, result2,comerr;
            result = "";
            int ci;
            FileStream fs = null;
            Encoding encoder = Encoding.UTF8;
            byte[] bytes;

           /* result = sp.ReadLine();
            MessageBox.Show(result);*/
            
            
            if ((model == "M1829")&&(dwxjudge==1))
            {
                string[] seq;
                string[] seq2;
                string ps;
                string ssresult;
                int xi;
                ci = 0;
          string areaxx,lapicxx;
          areaxx = "";
          lapicxx = "";
          testresult = "ok";
                do
            {
                comerr = "ok";
                    try
                {
                   /* if (ci != 0)
                    {
                        sp.Close();
                        sp.Open();

                    }*/
                    result = sp.ReadLine();
                //    MessageBox.Show(result);
                    if (result == "")
                    { MessageBox.Show(result); }
     

                        if ((result.IndexOf("O") != -1) && (teststatue == true) && (t17 != 99))
                    {
                        sp.WriteLine("S");
                    }
                    if (result.IndexOf("?") != -1) 
                    {
                        testresult = "ng";
                        this.Invoke((EventHandler)(delegate
                        {
                            this.TEST.Enabled = true;
                            //  dwxi = 1;
                        }

                                 ));
                        this.Invoke((EventHandler)(delegate
                        {
                            this.label1.ForeColor = Color.Red;
                            this.label1.Text = "重新测试";
                        }

                                                                    ));
                      
                        // sp.DiscardInBuffer();
                       // sp.DiscardOutBuffer();
                     //   MessageBox.Show(result);
                    //    comerr = "ng";
                      /*  Delay(2000);
                          sp.Close();
                        sp.Open();
                        if (sp.IsOpen)
                        {
                            sp.NewLine = "\r\n";
                            sp.WriteLine("S");
                        }*/
                 
                    }
                 
                    if ((result.IndexOf("G") != -1) && (result.IndexOf("N") == -1))
                    {
                        //    MessageBox.Show(result);
                      //  sp.Close();
                       // sp.Open();
                        sp.WriteLine("J");
                      //  testresult = "ok";
                        // dataGridView1.Rows[x].Cells[0].Value = SN;
                        // string[] seq;
                      /*  this.Invoke((EventHandler)(delegate
                        {
                            this.label1.ForeColor = Color.Green;
                            this.label1.Text = "PASS";
                        }

                                                                                  ));*/


                    }
                    if (result.IndexOf("N") != -1)
                    {
                        //sp.Close();
                       // sp.Open();
                        sp.WriteLine("J");
                        testresult = "ng";
                        this.Invoke((EventHandler)(delegate
                        {
                            this.TEST.Enabled = true;
                          //  dwxi = 1;
                        }

                                 ));
                        this.Invoke((EventHandler)(delegate
                        {
                            this.label1.ForeColor = Color.Red;
                            this.label1.Text = "FAIL";
                        }

                                                                    ));




                    }

                    if ((result.IndexOf("=") != -1)&&(result.IndexOf("A") == -1))
                    {
                        sp.Close();
                        sp.Open();
                        sp.NewLine = "\r\n";
                        sp.WriteLine("S");
                      
                        /*testresult = "ng";
                        this.Invoke((EventHandler)(delegate
                        {
                            this.TEST.Enabled = true;
                            //  dwxi = 1;
                        }

                                 ));
                        this.Invoke((EventHandler)(delegate
                        {
                            this.label1.ForeColor = Color.Red;
                            this.label1.Text = "ComErr ReTest";
                        }

                                                                    ));

                        */

                    }
                    
                    
                    
                    if ((result.IndexOf("A=") != -1)&&(result.IndexOf("L=") != -1))
                    {
                        // MessageBox.Show(result);
                       // sp.Close();
                     //   sp.WriteLine("ZB");
                        this.Invoke((EventHandler)(delegate
                        {
                           // this.TEST.Enabled = true;
                        }

                            ));

                        teststatue = false;
                        //     dataGridView1.Rows[x].Cells[0].Value = SN;
                      
                        xi = 0;
                       
                        ps = "";
                        seq2 = new string[5];
                        seq = result.Split('%');
                        foreach (string azu in seq)
                        {
                            seq2[xi] = azu;
                            xi++;

                        }

                        areaxx = areaxx + seq2[0].Substring(seq2[0].IndexOf("=") + 1, seq2[0].Length - seq2[0].IndexOf("=") - 1)+",";
                        ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("=") - 1);
                        // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("="));
                        ps = ps.Replace(",", "");
                        lapicxx = lapicxx + ps + ",";
                        if (ci == 0)
                        {
                            MessageBox.Show("TEST 3,4");
                            if (comerr == "ng")
                            {
                                MessageBox.Show("Reset Com"); 
                                sp.Close();
                                   try
                                   {

                                       sp.Open();

                                       //sp.Close();
                                       // MessageBox.Show("OK");
                                   }
                                   catch (InvalidOperationException eee)
                                   {
                                       //  MessageBox.Show("yidakai");
                                   }

                                   catch (Exception ee)
                                   {
                                       MessageBox.Show(ee.Message);
                                       comtrue = false;
                                       TEST.Enabled = false;
                                       this.label1.ForeColor = Color.Red;
                                       this.label1.Text = "Com Err";
                                   }
                            }
                             sp.NewLine = "\r\n";
                            sp.WriteLine("S");
                        }
                        if (ci == 1)
                        { 
                            MessageBox.Show("TEST 5,6");
                            if (comerr == "ng")
                            {
                                MessageBox.Show("Reset Com");
                                sp.Close();
                                try
                                {

                                    sp.Open();

                                    //sp.Close();
                                    // MessageBox.Show("OK");
                                }
                                catch (InvalidOperationException eee)
                                {
                                    //  MessageBox.Show("yidakai");
                                }

                                catch (Exception ee)
                                {
                                    MessageBox.Show(ee.Message);
                                    comtrue = false;
                                    TEST.Enabled = false;
                                    this.label1.ForeColor = Color.Red;
                                    this.label1.Text = "Com Err";
                                }
                            }
                            sp.NewLine = "\r\n";
                            sp.WriteLine("S");
                        }
                        ci++;
                     //   sp.WriteLine("ZB");
                   
                    }
                
                
                
                
                
                
                
                }


                catch (Exception ee)
                {
                    MessageBox.Show(ee.Message);
                    this.Invoke((EventHandler)(delegate
                    {
                        this.TEST.Enabled = true;
                    }

                                                 ));
                }



            }while((ci<3)&&(testresult=="ok"));
                sp.Close();
                ssresult = ""; 
                this.Invoke((EventHandler)(delegate
                {
                    this.dataGridView1.Rows.Add();
                    this.dataGridView1.Rows[x].Cells[0].Value = SN;
                    this.dataGridView1.Rows[x].Cells[1].Value = areaxx;
                    this.dataGridView1.Rows[x].Cells[2].Value = lapicxx;
                    if (testresult == "ok")
                    {
                        this.dataGridView1.Rows[x].Cells[3].Value = "PASS";
                        this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.White;
                        ssresult = "PASS";
                    }
                    if (testresult == "ng")
                    {
                        this.dataGridView1.Rows[x].Cells[3].Value = "FAIL";
                        this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.Red;
                        ssresult = "FAIL";
                        dwxi = 1;
                    }

                    dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                    this.dataGridView1.Rows[x].Selected = true;
                    

                }

                          ));




                x++;
                if (testresult == "ok")
                {
                    savedata = new FileStream(savepath, FileMode.Append);
                    sw3 = new StreamWriter(savedata);
                    sw3.Write("SN: " + SN + "   ");
                    sw3.Write("AREA: " + areaxx + "   ");
                    //sw3.Write("AREA:" + seq2[0].Substring(seq2[0].Length - seq2[0].IndexOf("=")) + "\r\n");
                    // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("=") - 1);
                    // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("="));
                    // ps = ps.Replace(",", "");
                    sw3.Write("LAPLC: " + lapicxx + "   ");
                    sw3.Write("RESULT: " + ssresult + "\r\n");
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                    if (itemprocess == testitemno - 1)
                    {
                        fs = File.OpenWrite("d:\\DWTEST\\sncheck.txt");
                        fs.Position = fs.Length;
                        bytes = encoder.GetBytes(SN + "\r\n");
                        fs.Write(bytes, 0, bytes.Length);
                        fs.Close();
                    }
                    if(itemprocess==testitemno-1)
                  { 
                      this.label1.ForeColor = Color.Green;
                    this.label1.Text = "PASS";
                    this.TEST.Enabled = true;
                   // this.TEST.Enabled = true;
                    if (checkBox1.Checked == true)
                    {
                        dwxi = 1;
                        TEST.PerformClick();
                    }
                  }
                    if (itemprocess != testitemno - 1)
                    {  dwxi = 1;}
              
                
                }
          
            
           
            
            
            
            
            
            
            }
            
            
            
  //    分割线      

            if ((model == "M0075")&&(dwxjudge==1))
            {
                string[] seq;
                string[] seq2;
                string ps;
                string ssresult;
                int xi;
                ci = 0;
                string areaxx, lapicxx;
                areaxx = "";
                lapicxx = "";
                testresult = "ok";
                do
                {
                    try
                    {
                        result = sp.ReadLine();
                        if ((result.IndexOf("O") != -1) && (teststatue == true)&&(t17 != 99))
                        {
                            sp.WriteLine("S");
                        }
                        if (result.IndexOf("?") != -1)
                        {
                            sp.Close();
                            sp.Open();
                            sp.NewLine = "\r\n";
                            sp.WriteLine("S");
                        }
                     
                        if ((result.IndexOf("G") != -1) && (result.IndexOf("N") == -1))
                        {
                            //    MessageBox.Show(result);
                            sp.WriteLine("J");
                            //  testresult = "ok";
                            // dataGridView1.Rows[x].Cells[0].Value = SN;
                            // string[] seq;
                            /*  this.Invoke((EventHandler)(delegate
                              {
                                  this.label1.ForeColor = Color.Green;
                                  this.label1.Text = "PASS";
                              }

                                                                                        ));*/


                        }
                        if (result.IndexOf("N") != -1)
                        {
                            sp.WriteLine("J");
                            testresult = "ng";
                            this.Invoke((EventHandler)(delegate
                            {
                                this.TEST.Enabled = true;
                            }

                                     ));
                            this.Invoke((EventHandler)(delegate
                            {
                                this.label1.ForeColor = Color.Red;
                                this.label1.Text = "FAIL";
                            }

                                                                        ));




                        }

                        if (result.IndexOf("A=") != -1)
                        {
                            // MessageBox.Show(result);
                           // sp.Close();
                            this.Invoke((EventHandler)(delegate
                            {
                                //this.TEST.Enabled = true;
                            }

                                ));

                            teststatue = false;
                            //     dataGridView1.Rows[x].Cells[0].Value = SN;

                            xi = 0;

                            ps = "";
                            seq2 = new string[5];
                            seq = result.Split('%');
                            foreach (string azu in seq)
                            {
                                seq2[xi] = azu;
                                xi++;

                            }

                            areaxx = areaxx + seq2[0].Substring(seq2[0].IndexOf("=") + 1, seq2[0].Length - seq2[0].IndexOf("=") - 1) + ",";
                            ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("=") - 1);
                            // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("="));
                            ps = ps.Replace(",", "");
                            lapicxx = lapicxx + ps + ",";
                            if (ci == 0)
                            { 
                                MessageBox.Show("TEST 3,4");
                                sp.NewLine = "\r\n";
                                sp.WriteLine("S");
                            }
                         
                            ci++;
                        }







                    }


                    catch (Exception ee)
                    {
                        MessageBox.Show(ee.Message);
                        this.Invoke((EventHandler)(delegate
                        {
                            this.TEST.Enabled = true;
                        }

                                                     ));
                    }



                } while ((ci < 2) && (testresult == "ok"));
                sp.Close();
                ssresult = "";
                this.Invoke((EventHandler)(delegate
                {
                    this.dataGridView1.Rows.Add();
                    this.dataGridView1.Rows[x].Cells[0].Value = SN;
                    this.dataGridView1.Rows[x].Cells[1].Value = areaxx;
                    this.dataGridView1.Rows[x].Cells[2].Value = lapicxx;
                    if (testresult == "ok")
                    {
                        this.dataGridView1.Rows[x].Cells[3].Value = "PASS";
                        this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.White;
                        ssresult = "PASS";
                    }
                    if (testresult == "ng")
                    {
                        this.dataGridView1.Rows[x].Cells[3].Value = "FAIL";
                        this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.Red;
                        ssresult = "FAIL";
                        dwxi = 1;
                    }

                    dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                    this.dataGridView1.Rows[x].Selected = true;


                }

                          ));




                x++;
                if (testresult == "ok")
                {
                    savedata = new FileStream(savepath, FileMode.Append);
                    sw3 = new StreamWriter(savedata); 
                    sw3.Write("SN: " + SN + "   ");
                    sw3.Write("AREA: " + areaxx + "   ");
                    //sw3.Write("AREA:" + seq2[0].Substring(seq2[0].Length - seq2[0].IndexOf("=")) + "\r\n");
                    // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("=") - 1);
                    // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("="));
                    // ps = ps.Replace(",", "");
                    sw3.Write("LAPLC: " + lapicxx + "   ");
                    sw3.Write("RESULT: " + ssresult + "\r\n");
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                    if (itemprocess == testitemno - 1)
                    {
                        fs = File.OpenWrite("d:\\DWTEST\\sncheck.txt");
                        fs.Position = fs.Length;
                        bytes = encoder.GetBytes(SN + "\r\n");
                        fs.Write(bytes, 0, bytes.Length);
                        fs.Close();
                    }
                    if (itemprocess == testitemno - 1)
                    {
                        this.label1.ForeColor = Color.Green;
                        this.label1.Text = "PASS";
                        this.TEST.Enabled = true;
                        if (checkBox1.Checked == true)
                        {
                            dwxi = 1;
                            TEST.PerformClick(); 
                        }
                    }
                    if (itemprocess != testitemno - 1)
                    { dwxi = 1; }
                }








            }

            //分割线        
          
            
            
            
            
            
            
            //  分割线           
            if (((model == "M1818") || (model == "M1817") || (model == "M1816") || (model == "M181734") || (model == "M1816SHORT") || (model == "M1877") || (model == "M187734") || (model == "M1877SHORT")) && (dwxjudge == 1))
            {
                try
                {
                    result = sp.ReadLine();
                 //   MessageBox.Show(result);
                    /*   if ((result.IndexOf("O") != -1) && (teststatue == true)&&(t17!=99))
                    {
                        sp.WriteLine("S");
                    }*/
                    if ((result.IndexOf("O") != -1) && (teststatue == true) && (t17 == 17))
                    {
                      //  sp.WriteLine("U");
                        sp.WriteLine("B,1817    .034");
                    }
                    if (result.IndexOf("?") != -1)
                    {
                        sp.Close();
                        sp.Open();
                        sp.NewLine = "\r\n";
                        sp.WriteLine("S");
                    }
                  
                    if ((result.IndexOf("G") != -1) && (result.IndexOf("N") == -1))
                    {
                        //    MessageBox.Show(result);
                        sp.WriteLine("J");
                        testresult = "ok";
                        // dataGridView1.Rows[x].Cells[0].Value = SN;
                        // string[] seq;
                        this.Invoke((EventHandler)(delegate
                        {
                            /*this.label1.ForeColor = Color.Green;
                            this.label1.Text = "PASS";*/
                        }

                                                                                   ));


                    }
                    if (result.IndexOf("N") != -1)
                    {
                        sp.WriteLine("J");
                        testresult = "ng";
                    //    MessageBox.Show("NG");
                        this.Invoke((EventHandler)(delegate
                        {
                            this.TEST.Enabled = true;
                        }

                                 ));
                        this.Invoke((EventHandler)(delegate
                        {
                            this.label1.ForeColor = Color.Red;
                            this.label1.Text = "FAIL";
                        }

                                                                    ));




                    }
                    if (result.IndexOf("A=") != -1)
                    {
                        // MessageBox.Show(result);
                  //      MessageBox.Show("A=");
                        sp.Close();
                        this.Invoke((EventHandler)(delegate
                        {
                         //   this.TEST.Enabled = true;
                        }

                            ));

                        teststatue = false;
                        //     dataGridView1.Rows[x].Cells[0].Value = SN;
                        string[] seq;
                        string[] seq2;
                        string ps;
                        string ssresult;
                        int xi;
                        xi = 0;
                        ssresult = "";
                        ps = "";
                        seq2 = new string[5];
                        seq = result.Split('%');
                        foreach (string azu in seq)
                        {
                            seq2[xi] = azu;
                            xi++;

                        }
                        this.Invoke((EventHandler)(delegate
                     {
                         this.dataGridView1.Rows.Add();
                         this.dataGridView1.Rows[x].Cells[0].Value = SN;
                         this.dataGridView1.Rows[x].Cells[1].Value = seq2[0];
                         this.dataGridView1.Rows[x].Cells[2].Value = seq2[1];
                         if (testresult == "ok")
                         {
                             this.dataGridView1.Rows[x].Cells[3].Value = "PASS";
                             this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.White;
                             ssresult = "PASS";
                         }
                         if (testresult == "ng")
                         {
                           //  MessageBox.Show("NG222");
                             this.dataGridView1.Rows[x].Cells[3].Value = "FAIL";
                             this.dataGridView1.Rows[x].DefaultCellStyle.BackColor = Color.Red;
                             ssresult = "FAIL";
                             dwxi = 1;
                         }

                         dataGridView1.CurrentCell = this.dataGridView1.Rows[x].Cells[0];
                         this.dataGridView1.Rows[x].Selected = true;


                     }

                        ));




                        x++;
                        if (testresult == "ok")
                        {
                            savedata = new FileStream(savepath, FileMode.Append);
                            sw3 = new StreamWriter(savedata);
                            if ((itemstr[itemprocess].IndexOf("34") != -1)||(itemstr[itemprocess].IndexOf("Short") != -1))
                            { sw3.Write("SN: " + SN + "-34" + "   "); }
                            if ((itemstr[itemprocess].IndexOf("34") == -1) || (itemstr[itemprocess].IndexOf("Short") == -1))
                            { sw3.Write("SN: " +SN+"   "); }
                            sw3.Write("AREA: " + seq2[0].Substring(seq2[0].IndexOf("=") + 1, seq2[0].Length - seq2[0].IndexOf("=") - 1) + "   ");
                            //sw3.Write("AREA:" + seq2[0].Substring(seq2[0].Length - seq2[0].IndexOf("=")) + "\r\n");
                            ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("=") - 1);
                            // ps = seq2[1].Substring(seq2[1].IndexOf("=") + 1, seq2[1].Length - seq2[1].IndexOf("="));
                            ps = ps.Replace(",", "");
                            sw3.Write("LAPLC: " + ps + "   ");
                            sw3.Write("RESULT: " + ssresult + "\r\n");
                            sw3.Flush();
                            //关闭流
                            sw3.Close();
                            savedata.Close();
                            if (itemprocess == testitemno-1 )
                            {
                                fs = File.OpenWrite("d:\\DWTEST\\sncheck.txt");
                                fs.Position = fs.Length;
                                bytes = encoder.GetBytes(SN + "\r\n");
                                fs.Write(bytes, 0, bytes.Length);
                                fs.Close();
                            }
                            
                            if (itemprocess == testitemno-1 )
                            {
                                this.label1.ForeColor = Color.Green;
                                this.label1.Text = "PASS";
                                this.TEST.Enabled = true;
                                if (checkBox1.Checked == true)
                                {
                                    dwxi = 1;
                                    TEST.PerformClick(); 
                                }
                            }
                            if (itemprocess != testitemno-1 )
                            { dwxi = 1; }
                        
                        }
                    
                    
                    
                    
                    }
                }

                catch (Exception ee)
                {
                    MessageBox.Show(ee.Message);
                    this.Invoke((EventHandler)(delegate
                    {
                        this.TEST.Enabled = true;
                    }

                                                 ));
                }

            }
         
         
         }

        private void SaveFile()
        {






        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
          /*  sw3.Flush();
            //关闭流
            sw3.Close();
            savedata.Close();*/
            if (hcard >= 0)
            { cardrealease(hcard); }
            System.Environment.Exit(0);
        
        
        }

        private void m1817ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if ((equ == "DWX"))
            {
                sp.Open();
                if (sp.IsOpen)
                {
                    // t17 = 17;
                    sp.NewLine = "\r\n";
                    sp.WriteLine("B,1817    .012");
                    //   sp.WriteLine("U");

                }
                sp.Close();
                t17 = 99;
                m1817ToolStripMenuItem.Checked = true;
                m1818ToolStripMenuItem.Checked = false;
                m1829ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m0075ToolStripMenuItem.Checked = false;
                noSNToolStripMenuItem.Checked = false;
                m181734ToolStripMenuItem.Checked = false;
                model = "M1817";
                path2 = DateTime.Now.ToString("yyyy");
                //   MessageBox.Show(savepath);
                savepath = "d:\\DWTEST";
                savepath = savepath + "\\" + model;
                if (!Directory.Exists(savepath))
                {

                    // Create the directory it does not exist.

                    Directory.CreateDirectory(savepath);

                }
                savepath = savepath + "\\" + path2 + ".txt";
                // MessageBox.Show(savepath);
                if (!File.Exists(savepath))
                {
                    savedata = new FileStream(savepath, FileMode.Create);
                    sw3 = new StreamWriter(savedata);
                    filenew = true;
                }
                else
                {
                    savedata = new FileStream(savepath, FileMode.Append);
                    sw3 = new StreamWriter(savedata);
                    filenew = false;
                    sw3.Write("\r\n" + "\r\n" + "\r\n");
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }

            }

            if ((equ == "TH2883"))
            {
                string rr;
                rr = "";
                /* sp.Open();
                if (sp.IsOpen)
                {
                    sp.NewLine = "\r\n";
                    sp.WriteLine("B,1817    .012");


                }
                sp.Close();*/
             //   myDmm.WriteString("MMEM:LOAD:STAT 1", true);
                myDmm.WriteString("IVOLT:VOLT 3800V", true);
                myDmm.WriteString("IVOLT:TIMP 5", true);
                myDmm.WriteString("IVOLT:EIMP 0", true);
                myDmm.WriteString("COMP:AREA ON", true);
                myDmm.WriteString("COMP:AREA:RANG 0,2196", true);
                myDmm.WriteString("COMP:AREA:DIFF 10", true);
                myDmm.WriteString("COMP:DIFF OFF", true);
                myDmm.WriteString("COMP:CORONA ON", true);
                myDmm.WriteString("COMP:CORONA:RANG 0,2196", true);
                myDmm.WriteString("COMP:CORONA:DIFF 250", true);
                myDmm.WriteString("COMP:PHAS OFF", true);
                myDmm.WriteString("TRIG:SOUR BUS", true);
                myDmm.WriteString("SWAVE:TRIG", true);
                Delay(5000);
                this.TEST.Enabled = false;

             /*   do
                {
                    rr = myDmm.ReadString();
                    //  MessageBox.Show(result);
                } while (rr.IndexOf("END") == -1);*/
                this.TEST.Enabled = true;
                t17 = 99;
                m1817ToolStripMenuItem.Checked = true;
                m1818ToolStripMenuItem.Checked = false;
                m1829ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m0075ToolStripMenuItem.Checked = false;
                noSNToolStripMenuItem.Checked = false;
                m181734ToolStripMenuItem.Checked = false;
                model = "M1817";
                path2 = DateTime.Now.ToString("yyyy");
                //   MessageBox.Show(savepath);
                savepath = "d:\\DWTEST";
                savepath = savepath + "\\" + model;
                if (!Directory.Exists(savepath))
                {

                    // Create the directory it does not exist.

                    Directory.CreateDirectory(savepath);

                }
                savepath = savepath + "\\" + path2 + ".txt";
                // MessageBox.Show(savepath);
                if (!File.Exists(savepath))
                {
                    savedata = new FileStream(savepath, FileMode.Create);
                    sw3 = new StreamWriter(savedata);
                    filenew = true;
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }
                else
                {
                    savedata = new FileStream(savepath, FileMode.Append);
                    sw3 = new StreamWriter(savedata);
                    filenew = false;
                    sw3.Write("\r\n" + "\r\n" + "\r\n");
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }






            }
        
        
        
        
        
        
        
        
        
        
        
        }

        private void m1829ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if ((equ == "DWX"))
            {
                sp.Open();
                if (sp.IsOpen)
                {
                    sp.NewLine = "\r\n";
                    //   sp.WriteLine("B,1817    .034");
                    sp.WriteLine("U");

                }
                sp.Close();
                m1817ToolStripMenuItem.Checked = false;
                m1818ToolStripMenuItem.Checked = false;
                m1829ToolStripMenuItem.Checked = true;
                m1816ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m0075ToolStripMenuItem.Checked = false;
                noSNToolStripMenuItem.Checked = false;
                m181734ToolStripMenuItem.Checked = false;
                model = "M1829";
                path2 = DateTime.Now.ToString("yyyy");
                //   MessageBox.Show(savepath);
                savepath = "d:\\DWTEST";
                savepath = savepath + "\\" + model;
                if (!Directory.Exists(savepath))
                {

                    // Create the directory it does not exist.

                    Directory.CreateDirectory(savepath);

                }
                savepath = savepath + "\\" + path2 + ".txt";
                // MessageBox.Show(savepath);
                if (!File.Exists(savepath))
                {
                    savedata = new FileStream(savepath, FileMode.Create);
                    sw3 = new StreamWriter(savedata);
                    filenew = true;
                }
                else
                {
                    savedata = new FileStream(savepath, FileMode.Append);
                    sw3 = new StreamWriter(savedata);
                    filenew = false;
                    sw3.Write("\r\n" + "\r\n" + "\r\n");
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }
            }

            if ((equ == "TH2883"))
            {
                string rr;
                rr = "";
               //    myDmm.WriteString("MMEM:LOAD:STAT 2", true);
            /*    myDmm.WriteString("IVOLT:VOLT 3200V", true);

                myDmm.WriteString("IVOLT:TIMP 5", true);
                myDmm.WriteString("IVOLT:EIMP 0", true);
                myDmm.WriteString("COMP:AREA ON", true);
                myDmm.WriteString("COMP:AREA:RANG 0,6000", true);
                myDmm.WriteString("COMP:AREA:DIFF 10", true);
                myDmm.WriteString("COMP:DIFF OFF", true);
                myDmm.WriteString("COMP:CORONA ON", true);
                myDmm.WriteString("COMP:CORONA:RANG 0,6000", true);
                myDmm.WriteString("COMP:CORONA:DIFF 250", true);
                myDmm.WriteString("COMP:PHAS OFF", true);
                myDmm.WriteString("TRIG:SOUR BUS", true);
                myDmm.WriteString("SWAVE:TRIG", true);*/
              //  this.TEST.Enabled = false;
            /*    do
                {
                    rr = myDmm.ReadString();
                    //  MessageBox.Show(result);
                } while (rr.IndexOf("END") == -1);*/
                //Delay(5000);
                this.TEST.Enabled = true;



                m1817ToolStripMenuItem.Checked = false;
                m1818ToolStripMenuItem.Checked = false;
                m1829ToolStripMenuItem.Checked = true;
                m1816ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m0075ToolStripMenuItem.Checked = false;
                noSNToolStripMenuItem.Checked = false;
                m181734ToolStripMenuItem.Checked = false;
                model = "M1829";
                path2 = DateTime.Now.ToString("yyyy");
                //   MessageBox.Show(savepath);
                savepath = "d:\\DWTEST";
                savepath = savepath + "\\" + model;
                if (!Directory.Exists(savepath))
                {

                    // Create the directory it does not exist.

                    Directory.CreateDirectory(savepath);

                }
                savepath = savepath + "\\" + path2 + ".txt";
                // MessageBox.Show(savepath);
                if (!File.Exists(savepath))
                {
                    savedata = new FileStream(savepath, FileMode.Create);
                    sw3 = new StreamWriter(savedata);
                    filenew = true;
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }
                else
                {
                    savedata = new FileStream(savepath, FileMode.Append);
                    sw3 = new StreamWriter(savedata);
                    filenew = false;
                    sw3.Write("\r\n" + "\r\n" + "\r\n");
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }








            }


            if (equ == "TH2839")
            {
                string rr;
                rr = "";
                myDmm.WriteString("VOLT 0.1V", true);
                myDmm.WriteString("FUNC:IMP CPD", true);
                myDmm.WriteString("FUNC:IMP:RANG:AUTO ON", true);
                myDmm.WriteString("TRIG:SOUR BUS", true);
               
                this.TEST.Enabled = true;



                m1817ToolStripMenuItem.Checked = false;
                m1818ToolStripMenuItem.Checked = false;
                m1829ToolStripMenuItem.Checked = true;
                m1816ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m0075ToolStripMenuItem.Checked = false;
                noSNToolStripMenuItem.Checked = false;
                m181734ToolStripMenuItem.Checked = false;
                model = "M1829";
                path2 = DateTime.Now.ToString("yyyy");
                //   MessageBox.Show(savepath);
                savepath = "d:\\DWTEST";
                savepath = savepath + "\\" + model;
                if (!Directory.Exists(savepath))
                {

                    // Create the directory it does not exist.

                    Directory.CreateDirectory(savepath);

                }
                savepath = savepath + "\\" + path2 + ".txt";
                // MessageBox.Show(savepath);
                if (!File.Exists(savepath))
                {
                    savedata = new FileStream(savepath, FileMode.Create);
                    sw3 = new StreamWriter(savedata);
                    filenew = true;
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }
                else
                {
                    savedata = new FileStream(savepath, FileMode.Append);
                    sw3 = new StreamWriter(savedata);
                    filenew = false;
                    sw3.Write("\r\n" + "\r\n" + "\r\n");
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }






            }
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        }

        private void m1818ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if ((equ == "DWX"))
            {
                sp.Open();
                if (sp.IsOpen)
                {
                    // t17 = 11;
                    sp.NewLine = "\r\n";
                    //   sp.WriteLine("B,1817    .034");
                    sp.WriteLine("U");

                }
                sp.Close();
                m1817ToolStripMenuItem.Checked = false;
                m1818ToolStripMenuItem.Checked = true;
                m1829ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m0075ToolStripMenuItem.Checked = false;
                noSNToolStripMenuItem.Checked = false;
                m181734ToolStripMenuItem.Checked = false;
                model = "M1818";
                path2 = DateTime.Now.ToString("yyyy");
                //   MessageBox.Show(savepath);
                savepath = "d:\\DWTEST";
                savepath = savepath + "\\" + model;
                if (!Directory.Exists(savepath))
                {

                    // Create the directory it does not exist.

                    Directory.CreateDirectory(savepath);

                }
                savepath = savepath + "\\" + path2 + ".txt";
                // MessageBox.Show(savepath);
                if (!File.Exists(savepath))
                {
                    savedata = new FileStream(savepath, FileMode.Create);
                    sw3 = new StreamWriter(savedata);
                    filenew = true;
                }
                else
                {
                    savedata = new FileStream(savepath, FileMode.Append);
                    sw3 = new StreamWriter(savedata);
                    filenew = false;
                    sw3.Write("\r\n" + "\r\n" + "\r\n");
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }
            }

            if ((equ == "TH2883"))
            {
                string rr;
                rr = "";
              //     myDmm.WriteString("MMEM:LOAD:STAT 1", true);
                myDmm.WriteString("IVOLT:VOLT 4000V", true);
                myDmm.WriteString("IVOLT:TIMP 5", true);
                myDmm.WriteString("IVOLT:EIMP 0", true);
                myDmm.WriteString("COMP:AREA ON", true);
                myDmm.WriteString("COMP:AREA:RANG 0,894", true);
                myDmm.WriteString("COMP:AREA:DIFF 10", true);
                myDmm.WriteString("COMP:DIFF OFF", true);
                myDmm.WriteString("COMP:CORONA ON", true);
                myDmm.WriteString("COMP:CORONA:RANG 0,894", true);
                myDmm.WriteString("COMP:CORONA:DIFF 250", true);
                myDmm.WriteString("COMP:PHAS OFF", true);
                myDmm.WriteString("TRIG:SOUR BUS", true);
                myDmm.WriteString("SWAVE:TRIG", true);
                this.TEST.Enabled = false;
              /*  do
                {
                    rr = myDmm.ReadString();
                    //  MessageBox.Show(result);
                } while (rr.IndexOf("END") == -1);*/
                Delay(5000);
                this.TEST.Enabled = true;





                m1817ToolStripMenuItem.Checked = false;
                m1818ToolStripMenuItem.Checked = true;
                m1829ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m0075ToolStripMenuItem.Checked = false;
                noSNToolStripMenuItem.Checked = false;
                m181734ToolStripMenuItem.Checked = false;
                model = "M1818";
                path2 = DateTime.Now.ToString("yyyy");
                //   MessageBox.Show(savepath);
                savepath = "d:\\DWTEST";
                savepath = savepath + "\\" + model;
                if (!Directory.Exists(savepath))
                {

                    // Create the directory it does not exist.

                    Directory.CreateDirectory(savepath);

                }
                savepath = savepath + "\\" + path2 + ".txt";
                // MessageBox.Show(savepath);
                if (!File.Exists(savepath))
                {
                    savedata = new FileStream(savepath, FileMode.Create);
                    sw3 = new StreamWriter(savedata);
                    filenew = true;
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }
                else
                {
                    savedata = new FileStream(savepath, FileMode.Append);
                    sw3 = new StreamWriter(savedata);
                    filenew = false;
                    sw3.Write("\r\n" + "\r\n" + "\r\n");
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }





            }

            if ((equ == "TH2839"))
            {
                string rr;
                rr = "";
                myDmm.WriteString("VOLT 0.1V", true);
                myDmm.WriteString("FUNC:IMP ZTD", true);
                myDmm.WriteString("FUNC:IMP:RANG:AUTO ON", true);
                myDmm.WriteString("TRIG:SOUR BUS", true);

                /* myDmm.WriteString("IVOLT:EIMP 0", true);
                myDmm.WriteString("COMP:AREA ON", true);
                myDmm.WriteString("COMP:AREA:RANG 0,894", true);
                myDmm.WriteString("COMP:AREA:DIFF 10", true);
                myDmm.WriteString("COMP:DIFF OFF", true);
                myDmm.WriteString("COMP:CORONA ON", true);
                myDmm.WriteString("COMP:CORONA:RANG 0,894", true);
                myDmm.WriteString("COMP:CORONA:DIFF 250", true);
                myDmm.WriteString("COMP:PHAS OFF", true);
                myDmm.WriteString("TRIG:SOUR BUS", true);
                myDmm.WriteString("SWAVE:TRIG", true);
                this.TEST.Enabled = false;
                do
                {
                    rr = myDmm.ReadString();
                    //  MessageBox.Show(result);
                } while (rr.IndexOf("END") == -1);*/
                this.TEST.Enabled = true;





                m1817ToolStripMenuItem.Checked = false;
                m1818ToolStripMenuItem.Checked = true;
                m1829ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m0075ToolStripMenuItem.Checked = false;
                noSNToolStripMenuItem.Checked = false;
                m181734ToolStripMenuItem.Checked = false;
                model = "M1818";
                path2 = DateTime.Now.ToString("yyyy");
                //   MessageBox.Show(savepath);
                savepath = "d:\\DWTEST";
                savepath = savepath + "\\" + model;
                if (!Directory.Exists(savepath))
                {

                    // Create the directory it does not exist.

                    Directory.CreateDirectory(savepath);

                }
                savepath = savepath + "\\" + path2 + ".txt";
                // MessageBox.Show(savepath);
                if (!File.Exists(savepath))
                {
                    savedata = new FileStream(savepath, FileMode.Create);
                    sw3 = new StreamWriter(savedata);
                    filenew = true;
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }
                else
                {
                    savedata = new FileStream(savepath, FileMode.Append);
                    sw3 = new StreamWriter(savedata);
                    filenew = false;
                    sw3.Write("\r\n" + "\r\n" + "\r\n");
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }





            }
        
        
        
        
      
        
        
        
        
        
        
        
        
        
        }

        private void m1816ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if ((equ == "DWX"))
            {
                sp.Open();
                if (sp.IsOpen)
                {
                    //  t17 = 11;
                    sp.NewLine = "\r\n";
                       sp.WriteLine("B,1816    .   ");
                    sp.WriteLine("U");

                }
                sp.Close();
                m1817ToolStripMenuItem.Checked = false;
                m1818ToolStripMenuItem.Checked = false;
                m1829ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = true;
                m1816SHORTToolStripMenuItem.Checked = false;
                m0075ToolStripMenuItem.Checked = false;
                noSNToolStripMenuItem.Checked = false;
                m181734ToolStripMenuItem.Checked = false;
                model = "M1816";
                path2 = DateTime.Now.ToString("yyyy");
                //   MessageBox.Show(savepath);
                savepath = "d:\\DWTEST";
                savepath = savepath + "\\" + model;
                if (!Directory.Exists(savepath))
                {

                    // Create the directory it does not exist.

                    Directory.CreateDirectory(savepath);

                }
                savepath = savepath + "\\" + path2 + ".txt";
                // MessageBox.Show(savepath);
                if (!File.Exists(savepath))
                {
                    savedata = new FileStream(savepath, FileMode.Create);
                    sw3 = new StreamWriter(savedata);
                    filenew = true;
                }
                else
                {
                    savedata = new FileStream(savepath, FileMode.Append);
                    sw3 = new StreamWriter(savedata);
                    filenew = false;
                    sw3.Write("\r\n" + "\r\n" + "\r\n");
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }

            }
            
            }

        private void m0075ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if ((equ == "DWX"))
            {
                m1817ToolStripMenuItem.Checked = false;
                m1818ToolStripMenuItem.Checked = false;
                m1829ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m0075ToolStripMenuItem.Checked = true;
                noSNToolStripMenuItem.Checked = false;
                m181734ToolStripMenuItem.Checked = false;
                model = "M0075";
                path2 = DateTime.Now.ToString("yyyy");
                //   MessageBox.Show(savepath);
                savepath = "d:\\DWTEST";
                savepath = savepath + "\\" + model;
                if (!Directory.Exists(savepath))
                {

                    // Create the directory it does not exist.

                    Directory.CreateDirectory(savepath);

                }
                savepath = savepath + "\\" + path2 + ".txt";
                // MessageBox.Show(savepath);
                if (!File.Exists(savepath))
                {
                    savedata = new FileStream(savepath, FileMode.Create);
                    sw3 = new StreamWriter(savedata);
                    filenew = true;
                }
                else
                {
                    savedata = new FileStream(savepath, FileMode.Append);
                    sw3 = new StreamWriter(savedata);
                    filenew = false;
                    sw3.Write("\r\n" + "\r\n" + "\r\n");
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }
            }

            if ((equ == "TH2883"))
            {

                string rr;
                rr = "";
            //    myDmm.WriteString("MMEM:LOAD:STAT 2", true);
                myDmm.WriteString("IVOLT:VOLT 3200V", true);

                myDmm.WriteString("IVOLT:TIMP 5", true);
                myDmm.WriteString("IVOLT:EIMP 0", true);
                myDmm.WriteString("COMP:AREA ON", true);
                myDmm.WriteString("COMP:AREA:RANG 0,6000", true);
                myDmm.WriteString("COMP:AREA:DIFF 10", true);
                myDmm.WriteString("COMP:DIFF OFF", true);
                myDmm.WriteString("COMP:CORONA ON", true);
                myDmm.WriteString("COMP:CORONA:RANG 0,6000", true);
                myDmm.WriteString("COMP:CORONA:DIFF 250", true);
                myDmm.WriteString("COMP:PHAS OFF", true);
                myDmm.WriteString("TRIG:SOUR BUS", true);
                myDmm.WriteString("SWAVE:TRIG", true);
                this.TEST.Enabled = false;
             /*   do
                {
                    rr = myDmm.ReadString();
                    //  MessageBox.Show(result);
                } while (rr.IndexOf("END") == -1);*/
                Delay(5000);
                this.TEST.Enabled = true;



                m1817ToolStripMenuItem.Checked = false;
                m1818ToolStripMenuItem.Checked = false;
                m1829ToolStripMenuItem.Checked = true;
                m1816ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m0075ToolStripMenuItem.Checked = false;
                noSNToolStripMenuItem.Checked = false;
                m181734ToolStripMenuItem.Checked = false;
                model = "M0075";
                path2 = DateTime.Now.ToString("yyyy");
                //   MessageBox.Show(savepath);
                savepath = "d:\\DWTEST";
                savepath = savepath + "\\" + model;
                if (!Directory.Exists(savepath))
                {

                    // Create the directory it does not exist.

                    Directory.CreateDirectory(savepath);

                }
                savepath = savepath + "\\" + path2 + ".txt";
                // MessageBox.Show(savepath);
                if (!File.Exists(savepath))
                {
                    savedata = new FileStream(savepath, FileMode.Create);
                    sw3 = new StreamWriter(savedata);
                    filenew = true;
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }
                else
                {
                    savedata = new FileStream(savepath, FileMode.Append);
                    sw3 = new StreamWriter(savedata);
                    filenew = false;
                    sw3.Write("\r\n" + "\r\n" + "\r\n");
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }







            }


            if (equ == "TH2839")
            {
                string rr;
                rr = "";
                myDmm.WriteString("VOLT 0.1V", true);
                myDmm.WriteString("FUNC:IMP CPD", true);
                myDmm.WriteString("FUNC:IMP:RANG:AUTO ON", true);
                myDmm.WriteString("TRIG:SOUR BUS", true);
                this.TEST.Enabled = true;
                m1817ToolStripMenuItem.Checked = false;
                m1818ToolStripMenuItem.Checked = false;
                m1829ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m0075ToolStripMenuItem.Checked = true;
                noSNToolStripMenuItem.Checked = false;
                m181734ToolStripMenuItem.Checked = false;
                model = "M0075";
                path2 = DateTime.Now.ToString("yyyy");
                //   MessageBox.Show(savepath);
                savepath = "d:\\DWTEST";
                savepath = savepath + "\\" + model;
                if (!Directory.Exists(savepath))
                {

                    // Create the directory it does not exist.

                    Directory.CreateDirectory(savepath);

                }
                savepath = savepath + "\\" + path2 + ".txt";
                // MessageBox.Show(savepath);
                if (!File.Exists(savepath))
                {
                    savedata = new FileStream(savepath, FileMode.Create);
                    sw3 = new StreamWriter(savedata);
                    filenew = true;
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }
                else
                {
                    savedata = new FileStream(savepath, FileMode.Append);
                    sw3 = new StreamWriter(savedata);
                    filenew = false;
                    sw3.Write("\r\n" + "\r\n" + "\r\n");
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }



            }
        
        
        
        
        
        
        }

        private void noSNToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m1817ToolStripMenuItem.Checked = false;
            m1818ToolStripMenuItem.Checked = false;
            m1829ToolStripMenuItem.Checked = false;
            m1816ToolStripMenuItem.Checked = false;
            m1816ToolStripMenuItem.Checked = false;
            m0075ToolStripMenuItem.Checked = false;
            noSNToolStripMenuItem.Checked = true;
            m181734ToolStripMenuItem.Checked = false;
            model = "NoSN";
            savepath = "d:\\DWTEST";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = savepath + "\\" + model;
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        private void add_Click(object sender, EventArgs e)
        {
            int actday;
            string daystr;
            daystr = "";
            actday = int.Parse(textBox1.Text);
            if (actday == 365)
            { actday = 1; }
            else
            { actday = actday + 1; }
            if (actday < 10)
            { daystr = "00" + actday.ToString(); }
            if ((actday >= 10) && (actday < 100))
            { daystr = "0" + actday.ToString(); }
            if (actday >= 100)
            { daystr = actday.ToString(); }
            textBox1.Text = daystr;
        }

        private void reduce_Click(object sender, EventArgs e)
        {
            int actday;
            string daystr;
            daystr = "";
            actday = int.Parse(textBox1.Text);
            if (actday == 1)
            { actday = 365; }
            else
            { actday = actday - 1; }
            if (actday < 10)
            { daystr = "00" + actday.ToString(); }
            if ((actday >= 10) && (actday < 100))
            { daystr = "0" + actday.ToString(); }
            if (actday >= 100)
            { daystr = actday.ToString(); }
            textBox1.Text = daystr;
        }

        private void m181734ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if ((equ == "DWX"))
            {
                sp.Open();
                if (sp.IsOpen)
                {
                    //  t17 = 17;
                    sp.NewLine = "\r\n";
                    sp.WriteLine("B,1817    .034");
                    // sp.WriteLine("U");

                }
                sp.Close();
                t17 = 99;
                m1817ToolStripMenuItem.Checked = false;
                m1818ToolStripMenuItem.Checked = false;
                m1829ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m0075ToolStripMenuItem.Checked = false;
                noSNToolStripMenuItem.Checked = false;
                m181734ToolStripMenuItem.Checked = true;
                model = "M181734";
                path2 = DateTime.Now.ToString("yyyy");
                //   MessageBox.Show(savepath);
                savepath = "d:\\DWTEST";
                savepath = savepath + "\\" + "M1817";
                if (!Directory.Exists(savepath))
                {

                    // Create the directory it does not exist.

                    Directory.CreateDirectory(savepath);

                }
                savepath = savepath + "\\" + path2 + ".txt";
                // MessageBox.Show(savepath);
                if (!File.Exists(savepath))
                {
                    savedata = new FileStream(savepath, FileMode.Create);
                    sw3 = new StreamWriter(savedata);
                    filenew = true;
                }
                else
                {
                    savedata = new FileStream(savepath, FileMode.Append);
                    sw3 = new StreamWriter(savedata);
                    filenew = false;
                    sw3.Write("\r\n" + "\r\n" + "\r\n");
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }

            }

            if ((equ == "TH2883"))
            {
                string rr;
                rr = "";
                /* sp.Open();
                if (sp.IsOpen)
                {
                    sp.NewLine = "\r\n";
                    sp.WriteLine("B,1817    .034");


                }
                sp.Close();*/
             //   myDmm.WriteString("MMEM:LOAD:STAT 3", true);
                myDmm.WriteString("IVOLT:VOLT 2000V", true);
                myDmm.WriteString("IVOLT:TIMP 5", true);
                myDmm.WriteString("IVOLT:EIMP 0", true);
                myDmm.WriteString("COMP:AREA ON", true);
                myDmm.WriteString("COMP:AREA:RANG 0,894", true);
                myDmm.WriteString("COMP:AREA:DIFF 10", true);
                myDmm.WriteString("COMP:DIFF OFF", true);
                myDmm.WriteString("COMP:CORONA ON", true);
                myDmm.WriteString("COMP:CORONA:RANG 0,894", true);
                myDmm.WriteString("COMP:CORONA:DIFF 250", true);
                myDmm.WriteString("COMP:PHAS OFF", true);
                myDmm.WriteString("TRIG:SOUR BUS", true);
                myDmm.WriteString("SWAVE:TRIG", true);
                this.TEST.Enabled = false;
                Delay(5000);
               /* do
                {
                    rr = myDmm.ReadString();
                    //  MessageBox.Show(result);
                } while (rr.IndexOf("END") == -1);*/
                this.TEST.Enabled = true;
                t17 = 99;
                m1817ToolStripMenuItem.Checked = false;
                m1818ToolStripMenuItem.Checked = false;
                m1829ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m0075ToolStripMenuItem.Checked = false;
                noSNToolStripMenuItem.Checked = false;
                m181734ToolStripMenuItem.Checked = true;
                model = "M181734";
                path2 = DateTime.Now.ToString("yyyy");
                //   MessageBox.Show(savepath);
                savepath = "d:\\DWTEST";
                savepath = savepath + "\\" + "M1817";
                if (!Directory.Exists(savepath))
                {

                    // Create the directory it does not exist.

                    Directory.CreateDirectory(savepath);

                }
                savepath = savepath + "\\" + path2 + ".txt";
                // MessageBox.Show(savepath);
                if (!File.Exists(savepath))
                {
                    savedata = new FileStream(savepath, FileMode.Create);
                    sw3 = new StreamWriter(savedata);
                    filenew = true;
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }
                else
                {
                    savedata = new FileStream(savepath, FileMode.Append);
                    sw3 = new StreamWriter(savedata);
                    filenew = false;
                    sw3.Write("\r\n" + "\r\n" + "\r\n");
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }







            }
        
        
        
        }

        private void fileCalToolStripMenuItem_Click(object sender, EventArgs e)
        {
        //    MessageBox.Show("");
          /*  FileStream fs2;      
            StreamReader sr3;
            string contect;
            string[] seq;
            string[] seq2;
            string[] seq3;
            string[] seq4;
            int ix, iy;
            ix = 2;
            iy = 2;
            Excel.Application app = new Excel.Application();
            Excel.Workbook wbook = app.Workbooks.Add(Type.Missing);
            Excel.Worksheet wsheet = (Excel.Worksheet)wbook.Worksheets[1];
            app.Visible = true;
            wsheet.Cells[1, 1] = "SN";
            wsheet.Cells[1, 2] = "Area Result";
            wsheet.Cells[1, 3] = "Area Result2";
            wsheet.Cells[1, 4] = "Area Result3";
            wsheet.Cells[1, 5] = "CP";
            wsheet.Cells[1, 6] = "SRF";
            fs2 = new FileStream(savepath, FileMode.Open);
           // fs2 = new FileStream("D:\\DWTEST\\M1817\\2018.txt", FileMode.Open);
            sr3 = new StreamReader(fs2);
            int xi;
        //    if ((model != "M1829") && ((equ == "DWX") || ((equ == "TH2883"))) && (model != "M0075"))
          if(equ=="TH2883")
            {
                while (!sr3.EndOfStream)
                {
                    contect = sr3.ReadLine();
                    if (contect.IndexOf("SN") != -1)
                    {

                        xi = 0;
                        seq2 = new string[5];
                        seq = contect.Split(':');
                        foreach (string azu in seq)
                        {
                            seq2[xi] = azu;
                            xi++;

                        }
                        wsheet.Cells[ix, 1] = seq2[1];
                        ix++;

                    }

                    if (contect.IndexOf("AREA") != -1)
                    {

                        xi = 0;
                        seq2 = new string[5];
                        seq = contect.Split(':');
                        foreach (string azu in seq)
                        {
                            seq2[xi] = azu;
                            xi++;

                        }
                        wsheet.Cells[iy, 2] = seq2[1];
                        iy++;

                    }










                    //  MessageBox.Show(comstr);

                }



                wsheet.SaveAs("d:\\DWTEST\\data.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            //  MessageBox.Show(comstr);app
          //  wsheet.SaveAs("d:\\DWTEST\\data.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
          if ((model != "M1817") && (equ == "DWX") && (model != "M181734"))
          {
              while (!sr3.EndOfStream)
              {
                  contect = sr3.ReadLine();
                  if (contect.IndexOf("SN") != -1)
                  {

                      xi = 0;
                      seq2 = new string[5];
                      seq = contect.Split(':');
                      foreach (string azu in seq)
                      {
                          seq2[xi] = azu;
                          xi++;

                      }
                      wsheet.Cells[ix, 1] = seq2[1];
                      ix++;

                  }

                  if (contect.IndexOf("AREA") != -1)
                  {

                      xi = 0;
                      seq2 = new string[5];
                      seq = contect.Split(':');
                      foreach (string azu in seq)
                      {
                          seq2[xi] = azu;
                          xi++;

                      }
                      wsheet.Cells[iy, 2] = seq2[1];
                      iy++;

                  }










                  //  MessageBox.Show(comstr);

              }


              wsheet.SaveAs("d:\\DWTEST\\data.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);


          }

          if (((model == "M1829") || (model == "M0075")) && (equ == "DWX"))
          {
              while (!sr3.EndOfStream)
              {
                  contect = sr3.ReadLine();
                  if (contect.IndexOf("SN") != -1)
                  {

                      xi = 0;
                      seq2 = new string[5];
                      seq = contect.Split(':');
                      foreach (string azu in seq)
                      {
                          seq2[xi] = azu;
                          xi++;

                      }
                      wsheet.Cells[ix, 1] = seq2[1];
                      ix++;

                  }

                  if (contect.IndexOf("AREA") != -1)
                  {

                      xi = 0;
                      seq2 = new string[5];
                      seq = contect.Split(':');
                      foreach (string azu in seq)
                      {
                          seq2[xi] = azu;
                          xi++;

                      }
                      xi = 0;
                      seq4 = new string[5];
                      seq3 = seq2[1].Split(',');
                      foreach (string azu2 in seq3)
                      {
                          seq4[xi] = azu2;
                          xi++;

                      }
                      int ll;
                      for (ll = 0; ll < xi; ll++)
                      {
                          wsheet.Cells[iy, ll+2] = seq4[ll];
                      }
                          iy++;
                      
                  }










                  //  MessageBox.Show(comstr);

              }


              wsheet.SaveAs("d:\\DWTEST\\data.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);


          }

          if((equ == "TH2839")&&(model=="M1818"))
          {
              while (!sr3.EndOfStream)
              {
                  contect = sr3.ReadLine();
                  if (contect.IndexOf("SN") != -1)
                  {

                      xi = 0;
                      seq2 = new string[5];
                      seq = contect.Split(':');
                      foreach (string azu in seq)
                      {
                          seq2[xi] = azu;
                          xi++;

                      }
                      wsheet.Cells[ix, 1] = seq2[1];
                      ix++;

                  }

                  if (contect.IndexOf("Hz") != -1)
                  {

                      xi = 0;
                      seq2 = new string[5];
                      seq = contect.Split(':');
                      foreach (string azu in seq)
                      {
                          seq2[xi] = azu;
                          xi++;

                      }
                      wsheet.Cells[iy,6] = seq2[1];
                      iy++;

                  }










                  //  MessageBox.Show(comstr);

              }



              wsheet.SaveAs("d:\\DWTEST\\data.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
          }


          if ((equ == "TH2839") && (model == "M1829") && (model == "M0075"))
          {
              while (!sr3.EndOfStream)
              {
                  contect = sr3.ReadLine();
                  if (contect.IndexOf("SN") != -1)
                  {

                      xi = 0;
                      seq2 = new string[5];
                      seq = contect.Split(':');
                      foreach (string azu in seq)
                      {
                          seq2[xi] = azu;
                          xi++;

                      }
                      wsheet.Cells[ix, 1] = seq2[1];
                      ix++;

                  }

                  if (contect.IndexOf("Cp") != -1)
                  {

                      xi = 0;
                      seq2 = new string[5];
                      seq = contect.Split(':');
                      foreach (string azu in seq)
                      {
                          seq2[xi] = azu;
                          xi++;

                      }
                      wsheet.Cells[iy, 5] = seq2[1];
                      iy++;

                  }










                  //  MessageBox.Show(comstr);

              }



              wsheet.SaveAs("d:\\DWTEST\\data.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
          }                  
            
            
            
            
         
            
            
            
            fs2.Close();  
        
        
        
       
        
        
        
        
        
        */
        
        }

        private void dWX10ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            equ = "DWX";
            dWX10ToolStripMenuItem.Checked = true;
            dwxjudge = 1;
            //  tH2883ToolStripMenuItem.Checked = false;
           // tH2839ToolStripMenuItem.Checked = false;
           // tH2829ToolStripMenuItem.Checked = false;
            sp.Open();
            if (sp.IsOpen)
            {
                //  t17 = 17;
                sp.NewLine = "\r\n";
                //   sp.WriteLine("B,1817    .034");
                sp.WriteLine("U");

            }
            sp.Close();
        
        }

        private void tH2883ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            equ = "TH2883";
          //  dWX10ToolStripMenuItem.Checked = false;
            thjudge = 1;
            tH2883ToolStripMenuItem.Checked = true;
           // tH2839ToolStripMenuItem.Checked = false;
           // tH2829ToolStripMenuItem.Checked = false;
            myDmm.IO = (IMessage)rm.Open(equgpib2883, AccessMode.NO_LOCK, 2000, ""); //Open up a handle to the DMM with a 2 second timeout
            myDmm.IO.Timeout = 3000; //You can also set your timeout by doing this command, sets to 3 seconds

            //First start off with a reset state
            //  myDmm.IO.Clear(); //Send a device clear first to stop any measurements in process
            myDmm.WriteString("*IDN?", true); //Get the IDN string                
            Delay(1000);
            string IDN = myDmm.ReadString();
            // MessageBox.Show(IDN);
            if (IDN.IndexOf("TH2883") != -1)
            {
                this.label1.Text = "Ready";
                myDmm.WriteString("DISP:PAGE MEAS", true);
                myDmm.WriteString("DISP:GRID ON", true);

            }
            else
            {
                this.label1.Text = "Communication Err";
                // this.TEST.Enabled = false;
            }
       
        
        
        
        
        
        }

        private void tH2839ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            equ = "TH2839";
            tH2839ToolStripMenuItem.Checked = true;
            thjudge = 1;
            //    dWX10ToolStripMenuItem.Checked = false;
         //   tH2883ToolStripMenuItem.Checked = false;
         //   tH2829ToolStripMenuItem.Checked = false;
            myDmm.IO = (IMessage)rm.Open(equgpib2883, AccessMode.NO_LOCK, 2000, ""); //Open up a handle to the DMM with a 2 second timeout
            myDmm.IO.Timeout = 3000; //You can also set your timeout by doing this command, sets to 3 seconds

            //First start off with a reset state
            //  myDmm.IO.Clear(); //Send a device clear first to stop any measurements in process
            myDmm.WriteString("*IDN?", true); //Get the IDN string                
            Delay(1000);
            string IDN = myDmm.ReadString();
            // MessageBox.Show(IDN);
            if (IDN.IndexOf("TH2839") != -1)
            {
                this.label1.Text = "Ready";
                myDmm.WriteString("DISP:PAGE MEAS", true);
                myDmm.WriteString("AMPL:ALC ON", true);

            }
            else
            {
                this.label1.Text = "Communication Err";
                // this.TEST.Enabled = false;
            }
       
        
        
        
        
        }

        private void m1816SHORTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if ((equ == "DWX"))
            {
                sp.Open();
                if (sp.IsOpen)
                {
                    //  t17 = 11;
                    sp.NewLine = "\r\n";
                       sp.WriteLine("B,1816    .034");
                    sp.WriteLine("U");

                }
                sp.Close();
                m1817ToolStripMenuItem.Checked = false;
                m1818ToolStripMenuItem.Checked = false;
                m1829ToolStripMenuItem.Checked = false;
                m1816ToolStripMenuItem.Checked = false;
                m1816SHORTToolStripMenuItem.Checked = true;
                m0075ToolStripMenuItem.Checked = false;
                noSNToolStripMenuItem.Checked = false;
                m181734ToolStripMenuItem.Checked = false;
              
                model = "M1816SHORT";
                path2 = DateTime.Now.ToString("yyyy");
                //   MessageBox.Show(savepath);
                savepath = "d:\\DWTEST";
                savepath = savepath + "\\" + model;
                if (!Directory.Exists(savepath))
                {

                    // Create the directory it does not exist.

                    Directory.CreateDirectory(savepath);

                }
                savepath = savepath + "\\" + path2 + ".txt";
                // MessageBox.Show(savepath);
                if (!File.Exists(savepath))
                {
                    savedata = new FileStream(savepath, FileMode.Create);
                    sw3 = new StreamWriter(savedata);
                    filenew = true;
                }
                else
                {
                    savedata = new FileStream(savepath, FileMode.Append);
                    sw3 = new StreamWriter(savedata);
                    filenew = false;
                    sw3.Write("\r\n" + "\r\n" + "\r\n");
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savedata.Close();
                }

            }
       
        
        
        
        }

        private void tH2829ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            equ = "TH2829";
           // tH2839ToolStripMenuItem.Checked = false;
         //   dWX10ToolStripMenuItem.Checked = false;
         //   tH2883ToolStripMenuItem.Checked = false;
            tH2829ToolStripMenuItem.Checked = true;
            myDmm.IO = (IMessage)rm.Open(DutAddr, AccessMode.NO_LOCK, 2000, ""); //Open up a handle to the DMM with a 2 second timeout
            myDmm.IO.Timeout = 3000; //You can also set your timeout by doing this command, sets to 3 seconds

            //First start off with a reset state
            //  myDmm.IO.Clear(); //Send a device clear first to stop any measurements in process
            myDmm.WriteString("*IDN?\n", true); //Get the IDN string                
            Delay(1000);
            string IDN = myDmm.ReadString();
            // MessageBox.Show(IDN);
            if (IDN.IndexOf("T") != -1)
            {
                this.label1.Text = "Ready";
                myDmm.WriteString("DISP:PAGE TSD\n", true);
              /*  myDmm.WriteString("AMPL:ALC ON", true);*/

            }
            else
            {
                this.label1.Text = "Communication Err";
                // this.TEST.Enabled = false;
            }
        }

        private void gPIBSetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GPSET gpset = new GPSET();
           // ff.TransfEvent += frm_TransfEvent;
            DialogResult ddr = gpset.ShowDialog();
        
        
       
        }

        private void dWX1817ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sp.Open();
            if (sp.IsOpen)
            {
                // t17 = 17;
                sp.NewLine = "\r\n";
                sp.WriteLine("B,1817    .012");
                //   sp.WriteLine("U");

            }
            sp.Close();
            t17 = 99;
            dWX1817ToolStripMenuItem.Checked = true;
            tH28831817ToolStripMenuItem1.Checked = false;
            tH28291817ToolStripMenuItem1.Checked = false;
            model = "M1817";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + model;
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        
        
        }

        private void tH28831817ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";
 
      /*      myDmm.WriteString("IVOLT:VOLT 3800V", true);
            myDmm.WriteString("IVOLT:TIMP 5", true);
            myDmm.WriteString("IVOLT:EIMP 0", true);
            myDmm.WriteString("COMP:AREA ON", true);
            myDmm.WriteString("COMP:AREA:RANG 0,2196", true);
            myDmm.WriteString("COMP:AREA:DIFF 10", true);
            myDmm.WriteString("COMP:DIFF OFF", true);
            myDmm.WriteString("COMP:CORONA ON", true);
            myDmm.WriteString("COMP:CORONA:RANG 0,2196", true);
            myDmm.WriteString("COMP:CORONA:DIFF 250", true);
            myDmm.WriteString("COMP:PHAS OFF", true);
            myDmm.WriteString("TRIG:SOUR BUS", true);
            myDmm.WriteString("SWAVE:TRIG", true);
            this.TEST.Enabled = false;
            Delay(5000);*/
         /*   do
            {
                rr = myDmm.ReadString();
                //  MessageBox.Show(result);
            } while (rr.IndexOf("END") == -1);*/
            this.TEST.Enabled = true;
            t17 = 99;
            dWX1817ToolStripMenuItem.Checked = false;
            tH28831817ToolStripMenuItem1.Checked = true;
            tH28291817ToolStripMenuItem1.Checked = false;
            model = "M1817";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + model;
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
      
        
        }

        private void dWX1829ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sp.Open();
            if (sp.IsOpen)
            {
                sp.NewLine = "\r\n";
                //   sp.WriteLine("B,1817    .034");
                sp.WriteLine("U");

            }
            sp.Close();
            dWX1829ToolStripMenuItem.Checked = true;
            tH28831829ToolStripMenuItem1.Checked = false;
            tH28391829ToolStripMenuItem1.Checked = false;
            tH28291829ToolStripMenuItem1.Checked = false;
            model = "M1829";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + model;
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
       
        }

        private void tH28831829ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";
            //    myDmm.WriteString("MMEM:LOAD:STAT 2", true);
       /*     myDmm.WriteString("IVOLT:VOLT 3200V", true);

            myDmm.WriteString("IVOLT:TIMP 5", true);
            myDmm.WriteString("IVOLT:EIMP 0", true);
            myDmm.WriteString("COMP:AREA ON", true);
            myDmm.WriteString("COMP:AREA:RANG 0,6000", true);
            myDmm.WriteString("COMP:AREA:DIFF 10", true);
            myDmm.WriteString("COMP:DIFF OFF", true);
            myDmm.WriteString("COMP:CORONA ON", true);
            myDmm.WriteString("COMP:CORONA:RANG 0,6000", true);
            myDmm.WriteString("COMP:CORONA:DIFF 250", true);
            myDmm.WriteString("COMP:PHAS OFF", true);
            myDmm.WriteString("TRIG:SOUR BUS", true);
            myDmm.WriteString("SWAVE:TRIG", true);*/
          //  this.TEST.Enabled = false;
            //Delay(5000);
            /*   do
            {
                rr = myDmm.ReadString();
                //  MessageBox.Show(result);
            } while (rr.IndexOf("END") == -1);*/
            this.TEST.Enabled = true;
            dWX1829ToolStripMenuItem.Checked = false;
            tH28831829ToolStripMenuItem1.Checked = true;
            tH28391829ToolStripMenuItem1.Checked = false;
            tH28291829ToolStripMenuItem1.Checked = false;


      
            model = "M1829";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + model;
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
       
        
        }

        private void tH28391829ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";
            myDmm.WriteString("VOLT 0.1V", true);
            myDmm.WriteString("FUNC:IMP CPD", true);
            myDmm.WriteString("FUNC:IMP:RANG:AUTO ON", true);
            myDmm.WriteString("TRIG:SOUR BUS", true);

            this.TEST.Enabled = true;
            dWX1829ToolStripMenuItem.Checked = false;
            tH28831829ToolStripMenuItem1.Checked = false;
            tH28391829ToolStripMenuItem1.Checked = true;
            tH28291829ToolStripMenuItem1.Checked = false;



            model = "M1829";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + model;
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
      
        
        }

        private void dWX1818ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sp.Open();
            if (sp.IsOpen)
            {
                // t17 = 11;
                sp.NewLine = "\r\n";
                //   sp.WriteLine("B,1817    .034");
                sp.WriteLine("U");

            }
            sp.Close();
            dWX1818ToolStripMenuItem.Checked = true;
           /* tH28831818ToolStripMenuItem1.Checked =false;
            tH28391818ToolStripMenuItem1.Checked = false;*/
            model = "M1818";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + model;
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        private void tH28831818ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";
            //     myDmm.WriteString("MMEM:LOAD:STAT 1", true);
          /*  myDmm.WriteString("IVOLT:VOLT 4000V", true);
            myDmm.WriteString("IVOLT:TIMP 5", true);
            myDmm.WriteString("IVOLT:EIMP 0", true);
            myDmm.WriteString("COMP:AREA ON", true);
            myDmm.WriteString("COMP:AREA:RANG 0,894", true);
            myDmm.WriteString("COMP:AREA:DIFF 10", true);
            myDmm.WriteString("COMP:DIFF OFF", true);
            myDmm.WriteString("COMP:CORONA ON", true);
            myDmm.WriteString("COMP:CORONA:RANG 0,894", true);
            myDmm.WriteString("COMP:CORONA:DIFF 250", true);
            myDmm.WriteString("COMP:PHAS OFF", true);
            myDmm.WriteString("TRIG:SOUR BUS", true);
            myDmm.WriteString("SWAVE:TRIG", true);
            this.TEST.Enabled = false;
            Delay(5000);*/
            /* do
            {
                rr = myDmm.ReadString();
                //  MessageBox.Show(result);
            } while (rr.IndexOf("END") == -1);*/
            this.TEST.Enabled = true;
         //   dWX1818ToolStripMenuItem.Checked = false;
            tH28831818ToolStripMenuItem1.Checked = true;
          //  tH28391818ToolStripMenuItem1.Checked = false;

            model = "M1818";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + model;
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }

      
        }

        private void tH28391818ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";
            myDmm.WriteString("VOLT 0.1V", true);
            myDmm.WriteString("FUNC:IMP ZTD", true);
            myDmm.WriteString("FUNC:IMP:RANG:AUTO ON", true);
            myDmm.WriteString("TRIG:SOUR BUS", true);

            this.TEST.Enabled = true;

         /*   dWX1818ToolStripMenuItem.Checked = false;
            tH28831818ToolStripMenuItem1.Checked = false;*/
            tH28391818ToolStripMenuItem1.Checked = true;




            model = "M1818";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + model;
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
       
        }

        private void dWX0075ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dWX0075ToolStripMenuItem.Checked = true;
            tH28830075ToolStripMenuItem1.Checked = true;
            tH28390075ToolStripMenuItem1.Checked = true;

            model = "M0075";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + model;
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        private void dWX1876ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sp.Open();
            if (sp.IsOpen)
            {
                //  t17 = 11;
                sp.NewLine = "\r\n";
            //    sp.WriteLine("B,1816    .012");
                sp.WriteLine("U");

            }
            sp.Close();
            dWX1876ToolStripMenuItem.Checked = true;
            tH28831876ToolStripMenuItem1.Checked = false;
            
            model = "M1816";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + "M1816";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        private void tH28831876ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";
            //     myDmm.WriteString("MMEM:LOAD:STAT 1", true);
            /*myDmm.WriteString("IVOLT:VOLT 4000V", true);
            myDmm.WriteString("IVOLT:TIMP 5", true);
            myDmm.WriteString("IVOLT:EIMP 0", true);
            myDmm.WriteString("COMP:AREA ON", true);
            myDmm.WriteString("COMP:AREA:RANG 0,894", true);
            myDmm.WriteString("COMP:AREA:DIFF 10", true);
            myDmm.WriteString("COMP:DIFF OFF", true);
            myDmm.WriteString("COMP:CORONA ON", true);
            myDmm.WriteString("COMP:CORONA:RANG 0,894", true);
            myDmm.WriteString("COMP:CORONA:DIFF 250", true);
            myDmm.WriteString("COMP:PHAS OFF", true);
            myDmm.WriteString("TRIG:SOUR BUS", true);
            myDmm.WriteString("SWAVE:TRIG", true);
            this.TEST.Enabled = false;
            Delay(5000);*/
            /*  do
            {
                rr = myDmm.ReadString();
                //  MessageBox.Show(result);
            } while (rr.IndexOf("END") == -1);*/
            this.TEST.Enabled = true;
            dWX1876ToolStripMenuItem.Checked = false;
            tH28831876ToolStripMenuItem1.Checked = true;

            model = "M1816";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + "M1816";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        private void tH28291876ToolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void tH28830075ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";
            //    myDmm.WriteString("MMEM:LOAD:STAT 2", true);
            myDmm.WriteString("IVOLT:VOLT 3200V", true);

            myDmm.WriteString("IVOLT:TIMP 5", true);
            myDmm.WriteString("IVOLT:EIMP 0", true);
            myDmm.WriteString("COMP:AREA ON", true);
            myDmm.WriteString("COMP:AREA:RANG 0,6000", true);
            myDmm.WriteString("COMP:AREA:DIFF 10", true);
            myDmm.WriteString("COMP:DIFF OFF", true);
            myDmm.WriteString("COMP:CORONA ON", true);
            myDmm.WriteString("COMP:CORONA:RANG 0,6000", true);
            myDmm.WriteString("COMP:CORONA:DIFF 250", true);
            myDmm.WriteString("COMP:PHAS OFF", true);
            myDmm.WriteString("TRIG:SOUR BUS", true);
            myDmm.WriteString("SWAVE:TRIG", true);
            this.TEST.Enabled = false;
            Delay(5000);
            /*    do
            {
                rr = myDmm.ReadString();
                //  MessageBox.Show(result);
            } while (rr.IndexOf("END") == -1);*/
            this.TEST.Enabled = true;

            tH28830075ToolStripMenuItem1.Checked = true;
            tH28390075ToolStripMenuItem1.Checked = false;
            dWX0075ToolStripMenuItem.Checked = false;
          
            model = "M0075";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + model;
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        private void tH28390075ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";
            myDmm.WriteString("VOLT 0.1V", true);
            myDmm.WriteString("FUNC:IMP CPD", true);
            myDmm.WriteString("FUNC:IMP:RANG:AUTO ON", true);
            myDmm.WriteString("TRIG:SOUR BUS", true);
            this.TEST.Enabled = true;
            tH28830075ToolStripMenuItem1.Checked = false;
            tH28390075ToolStripMenuItem1.Checked = true;
            dWX0075ToolStripMenuItem.Checked = false;
            model = "M0075";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + model;
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        private void dWX181734ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sp.Open();
            if (sp.IsOpen)
            {
                //  t17 = 17;
                sp.NewLine = "\r\n";
                sp.WriteLine("B,1817    .034");
                // sp.WriteLine("U");

            }
            sp.Close();
            t17 = 99;
            dWX181734ToolStripMenuItem.Checked = true;
            tH2883183734ToolStripMenuItem1.Checked = false;
            model = "M181734";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + "M1817";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        private void tH2883183734ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";

            //   myDmm.WriteString("MMEM:LOAD:STAT 3", true);
         /*   myDmm.WriteString("IVOLT:VOLT 2000V", true);
            myDmm.WriteString("IVOLT:TIMP 5", true);
            myDmm.WriteString("IVOLT:EIMP 0", true);
            myDmm.WriteString("COMP:AREA ON", true);
            myDmm.WriteString("COMP:AREA:RANG 0,894", true);
            myDmm.WriteString("COMP:AREA:DIFF 10", true);
            myDmm.WriteString("COMP:DIFF OFF", true);
            myDmm.WriteString("COMP:CORONA ON", true);
            myDmm.WriteString("COMP:CORONA:RANG 0,894", true);
            myDmm.WriteString("COMP:CORONA:DIFF 250", true);
            myDmm.WriteString("COMP:PHAS OFF", true);
            myDmm.WriteString("TRIG:SOUR BUS", true);
            myDmm.WriteString("SWAVE:TRIG", true);
            this.TEST.Enabled = false;
            Delay(5000);*/
         /*   do
            {
                rr = myDmm.ReadString();
                //  MessageBox.Show(result);
            } while (rr.IndexOf("END") == -1);*/
            this.TEST.Enabled = true;
            t17 = 99;
            dWX181734ToolStripMenuItem.Checked = false;
            tH2883183734ToolStripMenuItem1.Checked = true;
            model = "M181734";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + "M1817";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        private void dWX187634ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sp.Open();
            if (sp.IsOpen)
            {
                //  t17 = 11;
             /*   sp.NewLine = "\r\n";
                sp.WriteLine("B,1816    .034");
                sp.WriteLine("U");*/

            }
            sp.Close();
            dWX187634ToolStripMenuItem.Checked = true;
            tH2883187634ToolStripMenuItem1.Checked = false;

            model = "M1816SHORT";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + "M1816";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        void TestItem_TransfEvent(string value)
        {
            tistr = value;
        }
        void ItemNo_TransintfEvent(int valuet)
        {
            testitemno = valuet;
        } 
     
        private void itemSetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int xi,xt,xr;
            string[] seq,seq2,relay;
            relay = new string[3];
            TestItem Ti = new TestItem();
            Ti.TranstifEvent += TestItem_TransfEvent;
            Ti.TransnofEvent += ItemNo_TransintfEvent;
            
            DialogResult ddr = Ti.ShowDialog();
            xi = 0;          
            seq = tistr.Split(',');
            foreach (string azu in seq)
            {
                itemstr[xi] = azu;
                xi++;

            }
            xt = 0; xr = 0;
            for (xt = 0; xt < xi; xt++)
            {
                if (itemstr[xt].IndexOf("1818") != -1)
                {
                    xr = 0;
                    seq2 = itemstr[xt].Split('&');
                    foreach (string azu in seq2)
                    {
                        relay[xr] = azu;
                        xr++;

                    }
                    rvalue1818 = Convert.ToUInt32(relay[1], 16);
                     //  MessageBox.Show(Convert.ToString(rvalue1818, 10));
                }
                if ((itemstr[xt].IndexOf("1877") != -1)||(itemstr[xt].IndexOf("1876") != -1))
                {
                    xr = 0;
                    seq2 = itemstr[xt].Split('&');
                    foreach (string azu in seq2)
                    {
                        relay[xr] = azu;
                        xr++;

                    }
                    rvalue1877 = Convert.ToUInt32(relay[1], 16);
                    //  MessageBox.Show(Convert.ToString(rvalue1818, 10));
                }    
                
                if (itemstr[xt].IndexOf("1817_12") != -1)
                {
                    xr = 0;
                    seq2 = itemstr[xt].Split('&');
                    foreach (string azu in seq2)
                    {
                        relay[xr] = azu;
                        xr++;

                    }
                    rvalue1817_12 = Convert.ToUInt32(relay[1], 16);
                 //   MessageBox.Show(Convert.ToString(rvalue1817_12, 10));
                }
                if (itemstr[xt].IndexOf("1817_34") != -1)
                {
                    xr = 0;
                    seq2 = itemstr[xt].Split('&');
                    foreach (string azu in seq2)
                    {
                        relay[xr] = azu;
                        xr++;

                    }
                    rvalue1817_34 = Convert.ToUInt32(relay[1], 16);
                 //   MessageBox.Show(Convert.ToString(rvalue1817_34, 10));
                }
            
          
            }
            /*  MessageBox.Show(itemstr[0]);
             MessageBox.Show(itemstr[1]);*/
        }

        private void tH2883187634ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";
            //     myDmm.WriteString("MMEM:LOAD:STAT 1", true);
          /*  myDmm.WriteString("IVOLT:VOLT 3000V", true);
            myDmm.WriteString("IVOLT:TIMP 5", true);
            myDmm.WriteString("IVOLT:EIMP 0", true);
            myDmm.WriteString("COMP:AREA ON", true);
            myDmm.WriteString("COMP:AREA:RANG 0,894", true);
            myDmm.WriteString("COMP:AREA:DIFF 10", true);
            myDmm.WriteString("COMP:DIFF OFF", true);
            myDmm.WriteString("COMP:CORONA ON", true);
            myDmm.WriteString("COMP:CORONA:RANG 0,894", true);
            myDmm.WriteString("COMP:CORONA:DIFF 250", true);
            myDmm.WriteString("COMP:PHAS OFF", true);
            myDmm.WriteString("TRIG:SOUR BUS", true);
            myDmm.WriteString("SWAVE:TRIG", true);
            this.TEST.Enabled = false;
            Delay(5000);*/
            /*  do
            {
                rr = myDmm.ReadString();
                //  MessageBox.Show(result);
            } while (rr.IndexOf("END") == -1);*/
            this.TEST.Enabled = true;
            dWX1876ToolStripMenuItem.Checked = false;
            tH28831876ToolStripMenuItem1.Checked = true;

            model = "M181634";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + "M1816";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
       
        
        
        }

        private void tH28831877ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";
            //     myDmm.WriteString("MMEM:LOAD:STAT 1", true);
        /*    myDmm.WriteString("IVOLT:VOLT 4000V", true);
            myDmm.WriteString("IVOLT:TIMP 5", true);
            myDmm.WriteString("IVOLT:EIMP 0", true);
            myDmm.WriteString("COMP:AREA ON", true);
            myDmm.WriteString("COMP:AREA:RANG 0,894", true);
            myDmm.WriteString("COMP:AREA:DIFF 10", true);
            myDmm.WriteString("COMP:DIFF OFF", true);
            myDmm.WriteString("COMP:CORONA ON", true);
            myDmm.WriteString("COMP:CORONA:RANG 0,894", true);
            myDmm.WriteString("COMP:CORONA:DIFF 250", true);
            myDmm.WriteString("COMP:PHAS OFF", true);
            myDmm.WriteString("TRIG:SOUR BUS", true);
            myDmm.WriteString("SWAVE:TRIG", true);
            this.TEST.Enabled = false;
            Delay(5000);*/
            /*  do
            {
                rr = myDmm.ReadString();
                //  MessageBox.Show(result);
            } while (rr.IndexOf("END") == -1);*/
            this.TEST.Enabled = true;
            dWX1877ToolStripMenuItem.Checked = false;
            tH28831877ToolStripMenuItem1.Checked = true;

            model = "M1877";
            M1877Lx = "3.8";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + "M1877";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        private void dWX1877ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sp.Open();
            if (sp.IsOpen)
            {
                //  t17 = 11;
                sp.NewLine = "\r\n";
                //    sp.WriteLine("B,1816    .012");
                sp.WriteLine("U");

            }
            sp.Close();
            dWX1877ToolStripMenuItem.Checked = true;
            tH28831877ToolStripMenuItem1.Checked = false;

            model = "M1877";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + "M1877";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        private void tH2883187734ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";
            //     myDmm.WriteString("MMEM:LOAD:STAT 1", true);
          /*  myDmm.WriteString("IVOLT:VOLT 3000V", true);
            myDmm.WriteString("IVOLT:TIMP 5", true);
            myDmm.WriteString("IVOLT:EIMP 0", true);
            myDmm.WriteString("COMP:AREA ON", true);
            myDmm.WriteString("COMP:AREA:RANG 0,894", true);
            myDmm.WriteString("COMP:AREA:DIFF 10", true);
            myDmm.WriteString("COMP:DIFF OFF", true);
            myDmm.WriteString("COMP:CORONA ON", true);
            myDmm.WriteString("COMP:CORONA:RANG 0,894", true);
            myDmm.WriteString("COMP:CORONA:DIFF 250", true);
            myDmm.WriteString("COMP:PHAS OFF", true);
            myDmm.WriteString("TRIG:SOUR BUS", true);
            myDmm.WriteString("SWAVE:TRIG", true);
            this.TEST.Enabled = false;
            Delay(5000);*/
            /*  do
            {
                rr = myDmm.ReadString();
                //  MessageBox.Show(result);
            } while (rr.IndexOf("END") == -1);*/
            this.TEST.Enabled = true;
            dWX187734ToolStripMenuItem.Checked = false;
            tH2883187734ToolStripMenuItem1.Checked = true;

            model = "M1877SHORT";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + "M1877";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        private void dWX187734ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sp.Open();
            if (sp.IsOpen)
            {
                //  t17 = 11;
                /*   sp.NewLine = "\r\n";
                   sp.WriteLine("B,1816    .034");
                   sp.WriteLine("U");*/

            }
            sp.Close();
            dWX187734ToolStripMenuItem.Checked = true;
            tH2883187734ToolStripMenuItem1.Checked = false;

            model = "M1877SHORT";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + "M1877";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        void IOCard_TransfEvent(string value)
        {
            IOtype = value;
        }
        void IOcard_TransintfEvent(Int16 valuet)
        {
            hcard = valuet;
        } 
        private void iOTypeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IOCARD io = new IOCARD();
            io.TranscardfEvent += IOCard_TransfEvent;
            io.TranscardnofEvent += IOcard_TransintfEvent;
            DialogResult ddr = io.ShowDialog();
            //MessageBox.Show(IOtype);
        
        }

        private void iODebugToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IODEBUG debug = new IODEBUG();
            debug.CardSet = hcard;
            DialogResult ddr = debug.ShowDialog();
        
        }

        private void m1876calToolStripMenuItem_Click(object sender, EventArgs e)
        {
            card7296P1Arelayout(hcard, rvalue1877);
            myDmm.WriteString("MMEM:LOAD:STAT 6", true);
            myDmm.WriteString("IVOLT:VADJ ON", true); 
            myDmm.WriteString("IVOLT:VOLT 4000V", true);
                myDmm.WriteString("IVOLT:TIMP 5", true);
                myDmm.WriteString("IVOLT:EIMP 0", true);
                myDmm.WriteString("COMP:AREA ON", true);
                myDmm.WriteString("COMP:AREA:RANG 0,894", true);
                myDmm.WriteString("COMP:AREA:DIFF 10", true);
                myDmm.WriteString("COMP:DIFF OFF", true);
                myDmm.WriteString("COMP:CORONA ON", true);
                myDmm.WriteString("COMP:CORONA:RANG 0,894", true);
                myDmm.WriteString("COMP:CORONA:DIFF 250", true);
                myDmm.WriteString("COMP:PHAS OFF", true);
                myDmm.WriteString("TRIG:SOUR BUS", true);
                myDmm.WriteString("SWAVE:TRIG", true);
                this.TEST.Enabled = false;
                Delay(5000);
                this.TEST.Enabled = true;
                myDmm.WriteString("MMEM:DEL:STAT 6", true);  
            myDmm.WriteString("MMEM:STOR:STAT 6,\"1876_12\"", true);
        }

        private void m1876ShortcalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            card7296P1Arelayout(hcard, rvalue1877);
            myDmm.WriteString("MMEM:LOAD:STAT 7", true);
            myDmm.WriteString("IVOLT:VADJ ON", true); 
            myDmm.WriteString("IVOLT:VOLT 3000V", true);
            myDmm.WriteString("IVOLT:TIMP 5", true);
            myDmm.WriteString("IVOLT:EIMP 0", true);
            myDmm.WriteString("COMP:AREA ON", true);
            myDmm.WriteString("COMP:AREA:RANG 0,894", true);
            myDmm.WriteString("COMP:AREA:DIFF 10", true);
            myDmm.WriteString("COMP:DIFF OFF", true);
            myDmm.WriteString("COMP:CORONA ON", true);
            myDmm.WriteString("COMP:CORONA:RANG 0,894", true);
            myDmm.WriteString("COMP:CORONA:DIFF 250", true);
            myDmm.WriteString("COMP:PHAS OFF", true);
            myDmm.WriteString("TRIG:SOUR BUS", true);
            myDmm.WriteString("SWAVE:TRIG", true);
            this.TEST.Enabled = false;
            Delay(5000);
            this.TEST.Enabled = true;
            myDmm.WriteString("MMEM:DEL:STAT 7", true);  
            myDmm.WriteString("MMEM:STOR:STAT 7,\"1876_34\"", true);
        }

        private void tH283945ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";
        
            this.TEST.Enabled = true;
            dWX187745ToolStripMenuItem.Checked = false;
            tH283945ToolStripMenuItem1.Checked = true;

            model = "M1877";
            M1877Lx = "4.5";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + "M1877";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        private void tH283954ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";
           
            this.TEST.Enabled = true;
            dWX187745SHORTToolStripMenuItem1.Checked = false;
            tH283954ToolStripMenuItem1.Checked = true;

            model = "M1877SHORT";
            M1877Lx = "4.5";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + "M1877";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        
        
        }

        private void tH288349ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";

            this.TEST.Enabled = true;
            dWX49ToolStripMenuItem.Checked = false;
            tH288349ToolStripMenuItem1.Checked = true;

            model = "M1877";
            M1877Lx = "4.9";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + "M1877";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        private void tH288349SHORTToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";

            this.TEST.Enabled = true;
            dWX49SHORTToolStripMenuItem.Checked = false;
            tH288349SHORTToolStripMenuItem1.Checked = true;

            model = "M1877SHORT";
            M1877Lx = "4.9";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + "M1877";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        private void tH288353ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";

            this.TEST.Enabled = true;
            dWX53ToolStripMenuItem.Checked = false;
            tH288353ToolStripMenuItem1.Checked = true;

            model = "M1877";
            M1877Lx = "5.3";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + "M1877";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        
        }

        private void tH288353SHORTToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";

            this.TEST.Enabled = true;
            dWX53SHORTToolStripMenuItem.Checked = false;
            tH288353SHORTToolStripMenuItem1.Checked = true;

            model = "M1877SHORT";
            M1877Lx = "5.3";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + "M1877";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        private void tH288359ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";

            this.TEST.Enabled = true;
            dWX59ToolStripMenuItem.Checked = false;
            tH288359ToolStripMenuItem1.Checked = true;

            model = "M1877";
            M1877Lx = "5.9";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + "M1877";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }

        private void tH288359SHORTToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string rr;
            rr = "";

            this.TEST.Enabled = true;
            dWX59SHORTToolStripMenuItem.Checked = false;
            tH288359SHORTToolStripMenuItem1.Checked = true;

            model = "M1877SHORT";
            M1877Lx = "5.9";
            path2 = DateTime.Now.ToString("yyyy");
            //   MessageBox.Show(savepath);
            savepath = "d:\\DWTEST";
            savepath = savepath + "\\" + "M1877";
            if (!Directory.Exists(savepath))
            {

                // Create the directory it does not exist.

                Directory.CreateDirectory(savepath);

            }
            savepath = savepath + "\\" + path2 + ".txt";
            // MessageBox.Show(savepath);
            if (!File.Exists(savepath))
            {
                savedata = new FileStream(savepath, FileMode.Create);
                sw3 = new StreamWriter(savedata);
                filenew = true;
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
            else
            {
                savedata = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savedata);
                filenew = false;
                sw3.Write("\r\n" + "\r\n" + "\r\n");
                sw3.Flush();
                //关闭流
                sw3.Close();
                savedata.Close();
            }
        }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
   
    
    
    
    
    
    
    
    
    
    
    
    
    
    }
}
