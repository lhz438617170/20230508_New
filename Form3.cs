using CeBianLan.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;


using System.Timers;
using System.IO.Ports;
using NPOI.XWPF.UserModel;

namespace CeBianLan
{
    public partial class Form3 : Form
    {
        int numericValue;
        //private Settings settings = new Settings();
        private static System.Timers.Timer timer;
        private static ManagementEventWatcher watcher;
        private SerialPort serialPort;


        public Form3()
        {
            InitializeComponent();
            serialPort = new SerialPort("COM8", 19200);
            serialPort.DataReceived += new SerialDataReceivedEventHandler(DataReceivedHandler);
            serialPort.ErrorReceived += new SerialErrorReceivedEventHandler(serialPort_ErrorReceived);
            //serialPort.Open();
            try
            {
                serialPort.Open();
            }
            catch (Exception e)
            {
                MessageBox.Show("无法连接串口： " + e.Message);
            }
            timer = new System.Timers.Timer();
            timer.Interval = 1000;
            timer.Elapsed += timer_Tick;
            timer.Start();
            /*Console.WriteLine("按任意键停止...");
            Console.ReadKey();*/
            if (!serialPort.IsOpen)
            {
                serialPort.Close();
            }
            
            
        }

        private void timer_Tick(object sender, ElapsedEventArgs e)
        {
            string[] portNames = SerialPort.GetPortNames();
            
            //bool isOpen = serialPort.IsOpen;
            if (portNames.Contains("COM8"))
            {
                label1.Text = "串口已打开";
            }
            else
            {
                label1.Text = "关闭";
                
            }
        }

        private void serialPort_ErrorReceived(object sender, SerialErrorReceivedEventArgs e)
        {
           /* if (e.EventType == SerialError.RXOver || e.EventType == SerialError.Overrun || e.EventType == SerialError.RXParity || e.EventType == SerialError.Frame || e.EventType == SerialError.TXFull)
            {
                MessageBox.Show("串口连接已断开");
            }*/
        }

        private void DataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            string indata = sp.ReadLine();
            Console.WriteLine("接收到数据：" + indata);
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            /*textBox1.Text = settings.tb1;*/
            int.TryParse(textBox1.Text, out numericValue);

           
        }

        

        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {
           /* settings.tb1 = textBox1.Text;
            settings.Save();*/
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show(numericValue.ToString());
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (int.TryParse(textBox1.Text, out int value))
            {
                if (value>200)
                {
                    MessageBox.Show("最大输入200");
                    textBox1.Text = "";
                    //numericValue = 200 * 1000;
                }
                else
                {
                    numericValue = value * 1000;
                }
                
            }
        }
    }
}
