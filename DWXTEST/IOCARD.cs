using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;
using System.Runtime.InteropServices;
namespace DWXTEST
{
    public partial class IOCARD : Form
    {
        public IOCARD()
        {
            InitializeComponent();
        }
        FileStream fs7;
        StreamWriter sw7;
        StreamReader sr7;
        public delegate void TransfDelegate(String value);
        public delegate void TransintfDelegate(Int16 valuet);
        string configrelay, str,iostr;
        
        [DllImport(@"C:\Windows\System32\LPCI7250.dll", EntryPoint = "Init_Card7296", CallingConvention = CallingConvention.Cdecl)]
        public static extern Int16 Init7296();
        [DllImport(@"C:\Windows\System32\LPCI7250.dll", EntryPoint = "Card7296P1AOUTPUT", CallingConvention = CallingConvention.Cdecl)]
        public static extern void card7296configP1A(Int16 cardno);
        [DllImport(@"C:\Windows\System32\LPCI7250.dll", EntryPoint = "Card7296P1ARELAYOUTPUT", CallingConvention = CallingConvention.Cdecl)]
        public static extern void card7296P1Arelayout(Int16 cardno, UInt32 data);
        private void IOCARD_Load(object sender, EventArgs e)
        {
            int index;
            comtype.Items.Add("PCI7296");
            comtype.Items.Add("LPCI7250");
            configrelay = "d:\\relayconfig.txt";
            fs7 = new FileStream(configrelay, FileMode.Open);
            sr7 = new StreamReader(fs7);
            while (!sr7.EndOfStream)
            {
                str = sr7.ReadLine();
               
                //  MessageBox.Show(comstr);
            }
            fs7.Close();
            index = comtype.FindString(str); 
            
            
            
            comtype.SelectedIndex = index;
            iostr = comtype.SelectedItem.ToString();
        
        }
        public event TransfDelegate TranscardfEvent;
        public event TransintfDelegate TranscardnofEvent;
        private void button1_Click(object sender, EventArgs e)
        {
            Int16 card=-1;
            string ioss;
            ioss = "";
            fs7 = new FileStream(configrelay, FileMode.Truncate);
            sw7 = new StreamWriter(fs7);
            //  sr2 = new StreamReader(fs);
            sw7.Write(comtype.SelectedItem.ToString() + "\r\n");
            sw7.Flush();
            //关闭流
            sw7.Close();
            fs7.Close();
            if (iostr != comtype.SelectedItem.ToString())
            {
                ioss = comtype.SelectedItem.ToString();
                if (comtype.SelectedItem.ToString().IndexOf("7296") != -1)
                {
                    card = Init7296();
                    if (card < 0)
                    {
                        MessageBox.Show("I/O Card Init Fail=" + card.ToString());
                    }
                    else
                    {
                        TranscardnofEvent(card);
                        card7296configP1A(card);
                        card7296P1Arelayout(card, 0X00);
                    }
                }
                TranscardfEvent(ioss);
                
            
            }


            this.Close();
        }
    
    
    
    }
}
