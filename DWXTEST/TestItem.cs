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
    public partial class TestItem : Form
    {
        public TestItem()
        {
            InitializeComponent();
        }
        Dictionary<string, ComboBox> comboxList2 = new Dictionary<string, ComboBox>();
        public delegate void TransfDelegate(String value);
        public delegate void TransintfDelegate(int valuet);
        string configti,str;
        FileStream fs4;
        StreamWriter sw4;
        StreamReader sr4;
      
        private void TestItem_Load(object sender, EventArgs e)
        {
            int i;
            string[] p;
            int index, pi,no;
            index = 1;
            pi = 0;
            configti = "d:\\Itemconfig.txt";
            p = new string[6];
            foreach (Control item in this.Controls)
            {
                if (item.GetType() == typeof(System.Windows.Forms.ComboBox) && item.Name.StartsWith("comboBox"))
                    comboxList2.Add(item.Name, (ComboBox)item);

            }
            ItemNo.Items.Add ("1");
            ItemNo.Items.Add("2");
            for (i = 0; i < 2; i++)
            {
                string actual_test = "comboBox" + (i+1).ToString();
                comboxList2[actual_test].Items.Add("DWX_1817_12&0C");
                comboxList2[actual_test].Items.Add("DWX_1817_34&03");
                comboxList2[actual_test].Items.Add("TH2883_1817_12&0C");
                comboxList2[actual_test].Items.Add("TH2883_1817_34&03");
                comboxList2[actual_test].Items.Add("DWX_1829");
                comboxList2[actual_test].Items.Add("TH2883_1829");
                comboxList2[actual_test].Items.Add("TH2839_1829");
                comboxList2[actual_test].Items.Add("DWX_1818&0C");
                comboxList2[actual_test].Items.Add("TH2883_1818&0C");
                comboxList2[actual_test].Items.Add("TH2839_1818");
                comboxList2[actual_test].Items.Add("DWX_18764.5_Open&0C");
                comboxList2[actual_test].Items.Add("DWX_18764.5_Short&0C");
                comboxList2[actual_test].Items.Add("DWX_18764.9_Open&0C");
                comboxList2[actual_test].Items.Add("DWX_18764.9_Short&0C");
                comboxList2[actual_test].Items.Add("DWX_18763.8_Open&0C");
                comboxList2[actual_test].Items.Add("DWX_18763.8_Short&0C");
                comboxList2[actual_test].Items.Add("DWX_18765.3_Open&0C");
                comboxList2[actual_test].Items.Add("DWX_18765.3_Short&0C");
                comboxList2[actual_test].Items.Add("DWX_18765.9_Open&0C");
                comboxList2[actual_test].Items.Add("DWX_18765.9_Short&0C");
                comboxList2[actual_test].Items.Add("TH2883_18764.5_Open&0C");
                comboxList2[actual_test].Items.Add("TH2883_18764.5_Short&0C");
                comboxList2[actual_test].Items.Add("TH2883_18764.9_Open&0C");
                comboxList2[actual_test].Items.Add("TH2883_18764.9_Short&0C");
                comboxList2[actual_test].Items.Add("TH2883_18763.8_Open&0C");
                comboxList2[actual_test].Items.Add("TH2883_18763.8_Short&0C");
                comboxList2[actual_test].Items.Add("TH2883_18765.3_Open&0C");
                comboxList2[actual_test].Items.Add("TH2883_18765.3_Short&0C");
                comboxList2[actual_test].Items.Add("TH2883_18765.9_Open&0C");
                comboxList2[actual_test].Items.Add("TH2883_18765.9_Short&0C");
                
                
                comboxList2[actual_test].Items.Add("DWX_1877_Open&0C");
                comboxList2[actual_test].Items.Add("DWX_1877_Short&0C");
                comboxList2[actual_test].Items.Add("TH2883_1877_Open&0C");
                comboxList2[actual_test].Items.Add("TH2883_1877_Short&0C");
                comboxList2[actual_test].Items.Add("DWX_0075&0C");
                comboxList2[actual_test].Items.Add("TH2883_0075&0C");
                comboxList2[actual_test].Items.Add("TH2839_0075");
            }

            fs4 = new FileStream(configti, FileMode.Open);
            sr4 = new StreamReader(fs4);
            while (!sr4.EndOfStream)
            {
                str = sr4.ReadLine();
                p[pi] = str;
                pi = pi + 1;
                //  MessageBox.Show(comstr);
            }
            fs4.Close();
            no =int.Parse(p[0]);
            index = ItemNo.FindString(p[0]);
            ItemNo.SelectedIndex = index;
            for (i = 0; i < no; i++)
            {
                string actual_test = "comboBox" + (i+1).ToString();
                index = comboxList2[actual_test].FindString(p[i+1]);
                comboxList2[actual_test].SelectedIndex = index;

            }
        
        
        }


        public event TransfDelegate TranstifEvent;
        public event TransintfDelegate TransnofEvent;
        private void button1_Click(object sender, EventArgs e)
        {
            int i, pi,no;
            string[] p;
            string tstr;
           tstr="";
            pi = 0;
            p = new string[6];
            fs4 = new FileStream(configti, FileMode.Truncate);
            sw4 = new StreamWriter(fs4);
            //  sr2 = new StreamReader(fs);
            sw4.Write(ItemNo.SelectedItem.ToString() + "\r\n");
            no = int.Parse(ItemNo.SelectedItem.ToString());
            for (i = 0; i < no; i++)
            {
                string actual_test = "comboBox" + (i+1).ToString();
            //    MessageBox.Show(comboxList2[actual_test].Text.ToString());
                sw4.Write(comboxList2[actual_test].Text.ToString() + "\r\n");
                //comboxList2[actual_test].Items[comboxList2[actual_test].SelectedIndex].ToString()
            }


            sw4.Flush();
            //关闭流
            sw4.Close();
            fs4.Close();
            for (i = 0; i < no; i++)
            {
                string actual_test = "comboBox" + (i+1).ToString();
                if (i < 1)
                { tstr = tstr + comboxList2[actual_test].Text.ToString() + ","; }
                else
                { tstr = tstr + comboxList2[actual_test].Text.ToString(); }
            }
           // MessageBox.Show(tstr);    
            no =int.Parse( ItemNo.SelectedItem.ToString());
            TranstifEvent(tstr);
            TransnofEvent(no);
            this.Close();
        
        
        }  
    
    
    
    
    
    
    
    
    
    
    
    }
}
