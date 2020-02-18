using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;

namespace DWXTEST
{
    public partial class GPSET : Form
    {
        public GPSET()
        {
            InitializeComponent();
        }

        string gpconfig2;
        FileStream fs3;
        StreamWriter sw3;
        StreamReader sr3;
        Dictionary<string, ComboBox> gpList = new Dictionary<string, ComboBox>();
        private void GPSET_Load(object sender, EventArgs e)
        {
            int i,index,i2;
            string[] g2883;
            gpconfig2 = "d:\\gpconfig.txt";
            g2883 = new string[10];
            foreach (Control item in this.Controls)
            {
                if (item.GetType() == typeof(System.Windows.Forms.ComboBox) && item.Name.StartsWith("comboBox"))
                    gpList.Add(item.Name, (ComboBox)item);

            }
            for (i = 0; i < 2; i++)
            {
                string actual_test = "comboBox" + (i+1).ToString();
                for (i2 = 0; i2 < 4; i2++)
                {
                    
                    gpList[actual_test].Items.Add("ASRL" + (i2 + 1).ToString() + "::INSTR");
                }
            }

            for (i = 0; i < 2; i++)
            {
                string actual_test = "comboBox" + (i+1).ToString();
                gpList[actual_test].Items.Add("NoNe");
                gpList[actual_test].Items.Add("USB0::0X0471::0X2883::QF40900001::0::INSTR");
            }
         /*      for (i = 0; i < 4; i++)
            {
                comboBox1.Items.Add("ASRL" + (i + 1).ToString() + "::INSTR");


            }
            comboBox1.Items.Add("USB0::0X0471::0X2883::QF40900001::0::INSTR");
            for (i = 0; i < 4; i++)
            {
                comboBox2.Items.Add("ASRL" + (i + 1).ToString() + "::INSTR");


            }
            comboBox2.Items.Add("USB0::0X0471::0X2883::QF40900001::0::INSTR");
          */
            
            fs3 = new FileStream(gpconfig2, FileMode.Open);
            sr3 = new StreamReader(fs3);
            i = 0;
            while (!sr3.EndOfStream)
            {
                g2883[i] = sr3.ReadLine();

                i = i + 1;
                //  MessageBox.Show(comstr);
            }
            fs3.Close();
            for (i = 0; i < 2; i++)
            {
                string actual_test = "comboBox" + (i+1).ToString();
                index = gpList[actual_test].FindString(g2883[i]);
                gpList[actual_test].SelectedIndex = index;
            }
        
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int i = 0;
            fs3 = new FileStream(gpconfig2, FileMode.Truncate);
            sw3 = new StreamWriter(fs3);
         /*   for (i = 0; i < 2; i++)
            {
                string actual_test = "comboBox" + (i+1).ToString();
                sw3.Write(gpList[actual_test].SelectedItem.ToString() + "\r\n");
            }*/
            sw3.Write(comboBox1.SelectedItem.ToString() + "\r\n");    
            sw3.Flush();
            //关闭流
            sw3.Close();
            fs3.Close();
            this.Close();
        
        }
    }
}
