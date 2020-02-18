using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;




namespace DWXTEST
{
    public partial class ComConfig : Form
    {
        FileStream fs;
        StreamWriter sw2;
        StreamReader sr2;
        string configtxt, comstr;
        Dictionary<string, ComboBox> comboxList = new Dictionary<string, ComboBox>();
        bool testcon;
       public delegate void TransfDelegate();
        public ComConfig()
        {
            InitializeComponent();
        
        }

        private void ComConfig_Load(object sender, EventArgs e)
        {
            string[] ports = SerialPort.GetPortNames();
            string[] p;
            int index,pi,i;
            testcon = true;
            index = 1;
            pi = 0;
            configtxt = "d:\\dwconfig.txt";
            p = new string[6];
            foreach (Control item in this.Controls)
            {
                if (item.GetType() == typeof(System.Windows.Forms.ComboBox) && item.Name.StartsWith("comboBox"))
                   comboxList.Add(item.Name, (ComboBox)item);
           
            }

            equno.Items.Add("1");
            equno.Items.Add("2");
            
            
            
            foreach (string port in ports)
            {
                comboBox1.Items.Add(port);
            }
            comboBox2.Items.Add("9600");
            comboBox2.Items.Add("115200");
            comboBox3.Items.Add("8");
            comboBox3.Items.Add("7");
            comboBox4.Items.Add("even");
            comboBox4.Items.Add("odd");
            comboBox4.Items.Add("mark");
            comboBox4.Items.Add("none");
            comboBox4.Items.Add("space");
            comboBox5.Items.Add("1");
            comboBox5.Items.Add("2");
            comboBox6.Items.Add("none");
            comboBox6.Items.Add("send");
            comboBox6.Items.Add("xonxoff");
        
            fs = new FileStream(configtxt, FileMode.Open);
            sr2 = new StreamReader(fs);
            while (!sr2.EndOfStream)
            {
                comstr = sr2.ReadLine();
                p[pi] = comstr;
                pi = pi + 1;
                //  MessageBox.Show(comstr);
            }
             fs.Close();
            
            for (i = 1; i < 7; i++)
             {
                 string actual_test = "comboBox" + i.ToString();
                 index = comboxList[actual_test].FindString(p[i - 1]);
                 comboxList[actual_test].SelectedIndex = index;
            }
         //   index = comboBox1.FindString(comstr);
           // comboBox1.SelectedIndex = index;
       
        
        
        }

        public event TransfDelegate TransfEvent;
        private void button1_Click(object sender, EventArgs e)
        {
          //  Form1 ff =  (Form1)this.Owner;
            int i,pi;
            string[] p;
            SerialPort sp;
            pi = 0;
            p = new string[6];
            fs = new FileStream(configtxt, FileMode.Truncate);
            sw2 = new StreamWriter(fs);
          //  sr2 = new StreamReader(fs);
            for (i = 1; i < 7; i++)
            {
                string actual_test = "comboBox" + i.ToString();
                sw2.Write(comboxList[actual_test].SelectedItem.ToString() + "\r\n");
            }
           
      
            sw2.Flush();
            //关闭流
            sw2.Close();
            fs.Close();
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
            sp = new SerialPort(p[0]);
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
         try
            {
                sp.Open();
                sp.Close();
               // MessageBox.Show("OK");
            }
         catch (Exception ee)
         {
             MessageBox.Show(ee.Message);
           //  TEST.Enabled = false;
             testcon = false;
         }
         if (testcon == true)
         {
             TransfEvent();
             this.Close();
         }
            
            
           
        }
    
    
    
    
    
    
    
    
    
    
    
    
    }
}
