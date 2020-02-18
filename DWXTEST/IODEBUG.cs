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
    public partial class IODEBUG : Form
    {
        public IODEBUG()
        {
            InitializeComponent();
        }
        [DllImport(@"C:\Windows\System32\LPCI7250.dll", EntryPoint = "Card7296P1ARELAYOUTPUT", CallingConvention = CallingConvention.Cdecl)]
        public static extern void card7296P1Arelayout(Int16 cardno, UInt32 data);
       
        public static Int16 iocard;
        public Int16 CardSet
        {
            set
            {
                iocard = value;
            }
        }

        private void IODEBUG_Load(object sender, EventArgs e)
        {
            int i;
            for (i = 0; i < 256; i++)
            {
                comdata.Items.Add(i.ToString("X2"));

            }
           // MessageBox.Show(iocard.ToString());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            UInt32 value;
            value = Convert.ToUInt32(comdata.SelectedItem.ToString(),16);
         //   MessageBox.Show(Convert.ToString(value,10));
            card7296P1Arelayout(iocard, value);
        }
    
    
    
    
    }
}
