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
    public partial class SnInput : Form
    {

        public delegate void TransfDelegate(String value);
        public delegate void TransintfDelegate(int valuet);
        public static string models,dayys,equs;
        public string ModelSet
        {
            set
            {
                models = value;
            }
        }

        public string DaySet
        {
            set
            {
               dayys = value;
            }
        }

        public string EquSet
        {
            set
            {
                equs = value;
            }
        }      
        
        
        
        public SnInput()
        {
            InitializeComponent();
        }

        private void SnInput_Load(object sender, EventArgs e)
        {
           // MessageBox.Show(models);
            int x = (System.Windows.Forms.SystemInformation.WorkingArea.Width - this.Size.Width) / 2;
            int y = (System.Windows.Forms.SystemInformation.WorkingArea.Height - this.Size.Height) / 1;
            this.StartPosition = FormStartPosition.Manual;
            this.Location = (Point)new Size(x,y);
            snbox.Focus();
        }
        public event TransfDelegate TransfEvent;
        public event TransintfDelegate TransintfEvent;

        private void button1_Click(object sender, EventArgs e)
        {
            StreamReader check = new StreamReader("d:\\DWTEST\\sncheck.txt",Encoding.Default);
            FileStream fs = null;
            Encoding encoder = Encoding.UTF8;
            byte[] bytes;
            string line,snpass;
            int ss,snlength,index;
            ss = 0;
            line = "begin";
            snpass = "ok";
            snlength = snbox.Text.Length;
          //  MessageBox.Show(snlength.ToString());
            if ((models == "M1816") || (models == "M1817") || (models == "M181734") || (models == "M1816SHORT") || (models == "M181634") || (models.IndexOf("M1877") != -1))
            {
               /* if ((models == "M1816") || (models == "M1816SHORT"))
                {   MessageBox.Show("测试1-2"); }*/
                do
                {
                    line = check.ReadLine();
                    //MessageBox.Show(line);
                    if (line == snbox.Text)
                    {
                        ss = ss + 1;
                    }

                } while ((line != null) && (ss < 2));
                check.Close();
                if (equs != "TH2829")
                {
                    if (ss < 2)
                    {
                        /* fs = File.OpenWrite("d:\\DWTEST\\sncheck.txt");
                         fs.Position = fs.Length;
                         bytes = encoder.GetBytes(snbox.Text+"\r\n");
                         fs.Write(bytes,0,bytes.Length);
                         fs.Close();*/
                        if (models.IndexOf("M1877") == -1)
                        {
                            if (snlength != 32)
                            {


                                MessageBox.Show("长度不符");
                                snbox.Text = "";
                                snpass = "ng";


                                // break;
                            }
                        }
                        if (models.IndexOf("M1877") != -1)
                        {
                            if (snlength != 32)
                            {


                                MessageBox.Show("长度不符");
                                snbox.Text = "";
                                snpass = "ng";


                                // break;
                            }
                        }  
                        
                        /* if (snlength != 32)
                        {
                            if (models.IndexOf("M1877") != -1)
                            {
                                MessageBox.Show("长度不符");
                                snbox.Text = "";
                                snpass = "ng";
                            }

                            // break;
                        }*/
                        if ((snlength == 32) && ((models == "M1817") || (models == "M181734")))
                        {
                            index = snbox.Text.IndexOf("TCP");
                            if (snbox.Text.IndexOf("(P)1071833-00-H") == -1)
                            {
                                MessageBox.Show("机种名错误");
                                snbox.Text = "";
                                snpass = "ng";
                            }
                            if (snbox.Text.Substring(index + 5, 3) != dayys)
                            {
                                MessageBox.Show("日期错误");
                                MessageBox.Show(dayys);
                                MessageBox.Show(snbox.Text.Substring(index + 5, 3));
                                snbox.Text = "";
                                snpass = "ng";

                            }


                        }
                        if ((snlength == 32) && (models.IndexOf("M1877") != -1))
                        {
                            index = snbox.Text.IndexOf("TCP");
                            if (snbox.Text.IndexOf("(P)1138081-00-E") == -1)
                            {
                                MessageBox.Show("机种名错误");
                                snbox.Text = "";
                                snpass = "ng";
                            }
                            if (snbox.Text.Substring(index + 5, 3) != dayys)
                            {
                                MessageBox.Show("日期错误");
                                MessageBox.Show(dayys);
                                MessageBox.Show(snbox.Text.Substring(index + 5, 3));
                                snbox.Text = "";
                                snpass = "ng";

                            }


                        }
                        if ((snlength == 32) && (models.IndexOf("M1816") != -1))
                        {
                            index = snbox.Text.IndexOf("TCP");
                            if (snbox.Text.IndexOf("(P)1138081-00-C") == -1)
                            {
                                MessageBox.Show("机种名错误");
                                snbox.Text = "";
                                snpass = "ng";
                            }
                            if (snbox.Text.Substring(index + 5, 3) != dayys)
                            {
                                MessageBox.Show("日期错误");
                                MessageBox.Show(dayys);
                                MessageBox.Show(snbox.Text.Substring(index + 5, 3));
                                snbox.Text = "";
                                snpass = "ng";

                            }


                        }
                    
                        
                        
                        if (snpass == "ok")
                        {
                            TransfEvent(snbox.Text);
                            TransintfEvent(1);
                            this.Close();
                        }



                    }
                    else
                    {
                        MessageBox.Show("SN Duplicate");
                        snbox.Text = "";
                    }
                }
                if (equs == "TH2829")
                {
                    if (ss < 1)
                    {
                        /* fs = File.OpenWrite("d:\\DWTEST\\sncheck.txt");
                         fs.Position = fs.Length;
                         bytes = encoder.GetBytes(snbox.Text+"\r\n");
                         fs.Write(bytes,0,bytes.Length);
                         fs.Close();*/
                        if (snlength != 32)
                        {
                            MessageBox.Show("长度不符");
                            snbox.Text = "";
                            snpass = "ng";
                            // break;
                        }
                        if ((snlength == 32) && ((models == "M1817") || (models == "M181734")))
                        {
                            index = snbox.Text.IndexOf("TCP");
                            if (snbox.Text.IndexOf("(P)1071833-00-H") == -1)
                            {
                                MessageBox.Show("机种名错误");
                                snbox.Text = "";
                                snpass = "ng";
                            }
                            if (snbox.Text.Substring(index + 5, 3) != dayys)
                            {
                                MessageBox.Show("日期错误");
                                MessageBox.Show(dayys);
                                MessageBox.Show(snbox.Text.Substring(index + 5, 3));
                                snbox.Text = "";
                                snpass = "ng";

                            }


                        }


                        if (snpass == "ok")
                        {
                            TransfEvent(snbox.Text);
                            TransintfEvent(1);
                            this.Close();
                        }



                    }
                    else
                    {
                        MessageBox.Show("SN Duplicate");
                        snbox.Text = "";
                    }
                }  
            
           
            }
            else
            {
               /* if (equs != "TH2829")
                { MessageBox.Show("等待测试"); }
                if( (equs == "TH2829")&&(models == "M1818"))
                { MessageBox.Show("等待测试"); }*/
                do
                {
                    line = check.ReadLine();
                    if (line == snbox.Text)
                    {
                        ss = ss + 1;
                    }

                } while ((line != null) && (ss < 1));
                check.Close();


                if (ss < 1)
                {
                    /*fs = File.OpenWrite("d:\\DWTEST\\sncheck.txt");
                    fs.Position = fs.Length;
                    bytes = encoder.GetBytes(snbox.Text + "\r\n");
                    fs.Write(bytes, 0, bytes.Length);
                    fs.Close();*/
                    if (snlength != 32)
                    {
                        MessageBox.Show("长度不符");
                        snbox.Text = "";
                        snpass = "ng";
                        // break;
                    }
                   
                    if ((snlength == 32) && (models == "M1829"))
                    {
                        index = snbox.Text.IndexOf("TCP");
                        if (snbox.Text.IndexOf("(P)1071917-00-E") == -1)
                        {
                            MessageBox.Show("机种名错误");
                            snbox.Text = "";
                            snpass = "ng";
                        }
                        if (snbox.Text.Substring(index + 5, 3) != dayys) 
                        {
                            MessageBox.Show("日期错误");
                            snbox.Text = "";
                            snpass = "ng";

                        }


                    } 
                    
                       if ((snlength == 32) && (models == "M1818"))
                    {
                        index = snbox.Text.IndexOf("TCP");
                        if (snbox.Text.IndexOf("(P)1080536-00-F") == -1)
                        {
                            MessageBox.Show("机种名错误");
                            snbox.Text = "";
                            snpass = "ng";
                        }
                        if (snbox.Text.Substring(index + 5, 3) != dayys) 
                        {
                            MessageBox.Show("日期错误");
                            snbox.Text = "";
                            snpass = "ng";

                        }


                    }      
                    
                    
                   
                    
                    
                    
                    if (snpass == "ok")
                    {
                        TransfEvent(snbox.Text);
                        TransintfEvent(1);
                        this.Close();
                    }
                    
                }
                else
                {
                    MessageBox.Show("SN Duplicate");
                    snbox.Text = "";
                }


            }
        
        
        }

        private void button2_Click(object sender, EventArgs e)
        {

            TransintfEvent(2);
            this.Close();
        
        }

        private void SnInput_KeyDown(object sender, KeyEventArgs e)
        {
           /* if (e.KeyCode == Keys.Enter)
            {

                //如果还有keypress事件，不让此快捷键触发其事件可加一句代码
                MessageBox.Show("765");
                e.Handled = true;   //将Handled设置为true，指示已经处理过KeyPress事件
                button1.PerformClick();////执行单击confirm1的动作

            }*/
        }

        private void snbox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {

                //如果还有keypress事件，不让此快捷键触发其事件可加一句代码
             //   MessageBox.Show("765");
                e.Handled = true;   //将Handled设置为true，指示已经处理过KeyPress事件
                button1.PerformClick();////执行单击confirm1的动作

            }
        }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    }
}
