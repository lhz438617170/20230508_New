using CeBianLan.Properties;
using MaterialSkin;
using MaterialSkin.Controls;
using Modbus.Device;
using MySql.Data.MySqlClient;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XWPF.UserModel;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Windows.Media.TextFormatting;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;



namespace CeBianLan
{
    public partial class Form1 : MaterialForm
    {
        #region 数据库、串口参数、配置及全局变量
        private readonly MaterialSkinManager materialSkinManager;
        //变量类
        private infos infoss=new infos();
        //连接对象
        MySqlConnection conn = null;
        //语句执行对象
        MySqlCommand comm = null;
        //语句执行结果数据对象
        MySqlDataReader dr = null;
        string strConn = "";

        String serialPortName;
        SerialPort serialPort1 = new SerialPort();

        string info = "";
        string infos = "";
        int countts = 0;
        infos ifs=new infos();

        //当天日期
        DateTime today = DateTime.Now;
        
        private static IModbusMaster master;
        private static SerialPort port;
        //写线圈或写寄存器数组
        private bool[] coilsBuffer;
        private ushort[] registerBuffer;
        //功能码
        private string functionCode;
        //功能码序号
        private int functionOder;
        //参数(分别为从站地址,起始地址,长度)
        private byte slaveAddress;
        private ushort startAddress;
        private ushort numberOfPoints;
        //串口参数
        private string portName;
        private int baudRate;
        private Parity parity;
        private int dataBits;
        private StopBits stopBits;
        //自动测试标志位
        private bool AutoFlag = false;
        //获取当前时间
        private System.DateTime Current_time;

        //定时器初始化
        private System.Timers.Timer t = new System.Timers.Timer(1000);

        private const int WM_DEVICE_CHANGE = 0x219;            //设备改变           
        private const int DBT_DEVICEARRIVAL = 0x8000;          //设备插入
        private const int DBT_DEVICE_REMOVE_COMPLETE = 0x8004; //设备移除


        //radiobutton
        private int numberOfRecords = 0;
        Series series1 = new Series();
        Series series2 = new Series();
        Series series3 = new Series();
        Series series4 = new Series();
        Series series5 = new Series();
        Series series6 = new Series();

        string banbeninfo = "水体藻类荧光光谱在线分析仪操作软件V1.0";

        //折线图
        private List<int> XList = new List<int>();
        private List<int> YList = new List<int>();
        private Random randoms = new Random();
        #endregion
        


        public Form1()
        {
            System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
            timer.Interval = 1000; // 设置计时器间隔为10秒
            timer.Tick += Timer_Tick; // 注册计时器的Tick事件处理方法
            timer.Start(); // 启动计时器
            //窗体UI颜色设置
            InitializeComponent();
            materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.EnforceBackcolorOnAllComponents = true;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(
                       Primary.Blue600,
                       Primary.Blue800,
                       Primary.Blue300,
                       Accent.Red100,
                       TextShade.WHITE);

           
            //panel2.BackColor = Color.WhiteSmoke;
            strConn = "Database = hz_test;Server = localhost;Port = 3306;Password = root;UserID = root";
            conn = new MySqlConnection(strConn);
            serialPort1.DataReceived += new SerialDataReceivedEventHandler(serialPort1_DataReceived);//绑定事件
            //panel3.BackColor = Color.FromArgb(220, 220, 220);
            //panel4.BackColor = Color.FromArgb(220, 220, 220);
            panel5.BackColor = Color.FromArgb(220,220,220);

            // 设置日期选择器的事件处理程序
            dateTimePicker1.ValueChanged += new EventHandler(dateTimePicker1_ValueChanged);
            
        }



        //定时器，每秒刷新一次数据
        private void Timer_Tick(object sender, EventArgs e)
        {
            Thread thread = new Thread(DoSomeWork);
            thread.Start();
            //shuaxinzhexiantu();
            thread.Abort();
        }

        private void DoSomeWork()
        {
            Thread.Sleep(1000);
        }


        
        #region 显示水藻信息
        public void xxxinfo()
        {
            //打开数据库连接
            conn.Open();
            //查询语句
            comm = new MySqlCommand("select allyls,addres,dtimer,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,fo,fv,fm,fvfm,sigma,cn,wendu,dianya,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl from ain LIMIT 0,1", conn);
            dr = comm.ExecuteReader(); /*查询*/
            while (dr.Read())
            {
                label1.Text = dr.GetString("dtimer");
                label22.Text = dr.GetString("fvfm");
                label14.Text = dr.GetString("allswl");
                label48.Text = dr.GetString("allyls");
                textBox1.Text = dr.GetString("lanzao");
                textBox2.Text = dr.GetString("lvzao");
                textBox3.Text = dr.GetString("guizao");
                textBox4.Text = dr.GetString("jiazao");
                textBox5.Text = dr.GetString("yinzao");
                textBox15.Text = dr.GetString("fo");
                textBox14.Text = dr.GetString("fv");
                textBox13.Text = dr.GetString("fm");
                textBox12.Text = dr.GetString("sigma");
                textBox11.Text = dr.GetString("cn");
                textBox19.Text = dr.GetString("zhuodu");
                textBox18.Text = dr.GetString("cdom");
                textBox17.Text = dr.GetString("dianya");
                textBox16.Text = dr.GetString("wendu");
                textBox10.Text = dr.GetString("lanswl");
                textBox9.Text = dr.GetString("lvswl");
                textBox8.Text = dr.GetString("guiswl");
                textBox7.Text = dr.GetString("jiaswl");
                textBox6.Text = dr.GetString("yinswl");


            }
            dr.Close();
            conn.Close();
        }
        #endregion


        
        #region 显示下拉框地址方法
        public void addresinfo()
        {
            //打开数据库连接
            conn.Open();
            //查询语句
            comm = new MySqlCommand("select DISTINCT addres from ain", conn);
            comboBox2.Text = "请选择地址";
            dr = comm.ExecuteReader(); /*查询*/

            while (dr.Read())
            {
                //把地址赋值到下拉框
                comboBox2.Items.Add(dr["addres".ToString()]);

            }
            dr.Close();
            conn.Close();
        }
        #endregion

        
        #region 显示折线图方法
        public void chartinfo()
        {
            // 获取 Chart 控件的 X 轴
            Axis sxAxis = chart1.ChartAreas[0].AxisX;
            // 将 X 轴的 Minimum 属性设置为 0
            sxAxis.Minimum = 0;
            //修改折线图数据
            chart1.ChartAreas[0].AxisX.MajorGrid.LineDashStyle = ChartDashStyle.NotSet;
            chart1.ChartAreas[0].AxisY.MajorGrid.LineDashStyle = ChartDashStyle.DashDotDot; //设置网格类型为虚线

            // 获取折线图的 Y 轴对象
            var yAxis = chart1.ChartAreas[0].AxisY;
            // 获取 Y 轴的刻度线对象，并设置其 LabelForeColor 属性为红色
            yAxis.MajorTickMark.LineColor = Color.Black;
            yAxis.LabelStyle.ForeColor = Color.Black;

            // 获取折线图的 X 轴对象
            var xAxis = chart1.ChartAreas[0].AxisX;
            // 获取 X 轴的刻度线对象，并设置其 LabelForeColor 属性为红色
            xAxis.MajorTickMark.LineColor = Color.Black;
            xAxis.LabelStyle.ForeColor = Color.Black;
            chart1.ChartAreas[0].AxisX.Minimum = 1;
            //折线图获取数据库值
            conn.Open();
            comm = new MySqlCommand("select allyls,lanzao,lvzao,guizao,jiazao,yinzao from ain limit 10", conn);
            dr = comm.ExecuteReader(); /*查询*/

            // 添加折线

            // 添加折线
            //标记点边框颜色      
            series1.MarkerBorderColor = Color.Orange;
            //标记点边框大小
            series1.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series1.MarkerColor = Color.Orange;//AxisColor
            //标记点大小
            series1.MarkerSize = 8;
            //标记点类型     
            series1.MarkerStyle = MarkerStyle.Circle;
            series1.ChartType = SeriesChartType.Line;
            series1.Color = Color.Orange;
            series1.BorderWidth = 2;
            series1.IsValueShownAsLabel = false;
            series1.Name = "总叶绿素";
            //Series series2 = new Series();
            //标记点边框颜色      
            series2.MarkerBorderColor = Color.Blue;
            //标记点边框大小
            series2.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series2.MarkerColor = Color.Blue;//AxisColor
            //标记点大小
            series2.MarkerSize = 8;
            //标记点类型     
            series2.MarkerStyle = MarkerStyle.Circle;
            series2.ChartType = SeriesChartType.Line;
            series2.Color = Color.Blue;
            series2.BorderWidth = 2;
            series2.IsValueShownAsLabel = false;
            series2.Name = "蓝藻";
            //Series series3 = new Series();
            //标记点边框颜色      
            series3.MarkerBorderColor = Color.Green;
            //标记点边框大小
            series3.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series3.MarkerColor = Color.Green;//AxisColor
            //标记点大小
            series3.MarkerSize = 8;
            //标记点类型     
            series3.MarkerStyle = MarkerStyle.Circle;
            series3.ChartType = SeriesChartType.Line;
            series3.Color = Color.Green;
            series3.BorderWidth = 2;
            series3.IsValueShownAsLabel = false;
            series3.Name = "绿藻";
            //Series series4 = new Series();
            //标记点边框颜色      
            series4.MarkerBorderColor = Color.Gray;
            //标记点边框大小
            series4.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series4.MarkerColor = Color.Gray;//AxisColor
            //标记点大小
            series4.MarkerSize = 8;
            //标记点类型     
            series4.MarkerStyle = MarkerStyle.Circle;
            series4.ChartType = SeriesChartType.Line;
            series4.Color = Color.Gray;
            series4.BorderWidth = 2;
            series4.IsValueShownAsLabel = false;
            series4.Name = "硅藻";
            //Series series5 = new Series();
            //标记点边框颜色      
            series5.MarkerBorderColor = Color.Red;
            //标记点边框大小
            series5.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series5.MarkerColor = Color.Red;//AxisColor
            //标记点大小
            series5.MarkerSize = 8;
            //标记点类型     
            series5.MarkerStyle = MarkerStyle.Circle;
            series5.ChartType = SeriesChartType.Line;
            series5.Color = Color.Red;
            series5.BorderWidth = 2;
            series5.IsValueShownAsLabel = false;
            series5.Name = "甲藻";
            //Series series6 = new Series();
            //标记点边框颜色      
            series6.MarkerBorderColor = Color.Pink;
            //标记点边框大小
            series6.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series6.MarkerColor = Color.Pink;//AxisColor
            //标记点大小
            series6.MarkerSize = 8;
            //标记点类型     
            series6.MarkerStyle = MarkerStyle.Circle;
            series6.Color = Color.Pink;
            series6.BorderWidth = 2;
            series6.IsValueShownAsLabel = false;
            series6.ChartType = SeriesChartType.Line;
            series6.Name = "隐藻";

            chart1.Series.Add(series1);
            chart1.Series.Add(series2);
            chart1.Series.Add(series3);
            chart1.Series.Add(series4);
            chart1.Series.Add(series5);
            chart1.Series.Add(series6);

            // 添加数据点
            int i = 0;
            while (dr.Read())
            {
                series1.Points.AddXY(i, dr.GetDecimal("allyls"));
                series2.Points.AddXY(i, dr.GetDecimal("lanzao"));
                series3.Points.AddXY(i, dr.GetDecimal("lvzao"));
                series4.Points.AddXY(i, dr.GetDecimal("guizao"));
                series5.Points.AddXY(i, dr.GetDecimal("jiazao"));
                series6.Points.AddXY(i, dr.GetDecimal("yinzao"));
                i++;
            }

            dr.Close();
            conn.Close();
        }
        #endregion

        
        #region 获取打开串口方法
        public void serportinfo()
        {
            //获取打开串口
            string[] ports = System.IO.Ports.SerialPort.GetPortNames();//获取电脑上可用串口号
            comboBox1.Items.AddRange(ports);//给comboBox1添加数据
            comboBox1.SelectedIndex = comboBox1.Items.Count > 0 ? 0 : -1;//如果里面有数据,显示第0个


            comboBox3.Text = "19200";/*默认波特率:115200*/
            comboBox4.Text = "1";/*默认停止位:1*/
            comboBox5.Text = "8";/*默认数据位:8*/
            comboBox6.Text = "无";/*默认奇偶校验位:无*/
        }
        #endregion


        private void Form1_Load(object sender, EventArgs e)
        {

            
            try
            {
                materialLabel5.Text = banbeninfo;

                //显示查询信息到textbox框
                xxxinfo();
                //显示地址到下拉框
                addresinfo();
                //显示折线图
                chartinfo();
                //获取串口打开
                serportinfo();
                xxinfoseripot();
            }
            catch (Exception) {; }
            
        }


       

        /// <summary>
        /// 清空日志
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            
        }
        #region  无


        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void txt_slave1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void txt_startAddr1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void txt_length_TextChanged(object sender, EventArgs e)
        {

        }

        private void label43_Click(object sender, EventArgs e)
        {

        }

        static void MyThread()
        {
            Thread.Sleep(100);
        }


        private void button1_Click_1(object sender, EventArgs e)
        {
            
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            
        }

        private void button_AutomaticTest_Click_1(object sender, EventArgs e)
        {
            
        }

        private void button_ClosePort_Click_1(object sender, EventArgs e)
        {
            
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            
        }

        
        #endregion

        /// <summary>
        /// 导出报表为Csv
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <param name="strFilePath">物理路径</param>
        /// <param name="tableheader">表头</param>
        /// <param name="columname">字段标题,逗号分隔</param>
        public static bool dt2csv(DataTable dt, string strFilePath, string tableheader, string columname)
        {
            try
            {
                string strBufferLine = "";
                StreamWriter strmWriterObj = new StreamWriter(strFilePath, false, System.Text.Encoding.UTF8);
                strmWriterObj.WriteLine(tableheader);
                strmWriterObj.WriteLine(columname);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    strBufferLine = "";
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (j > 0)
                            strBufferLine += ",";
                        strBufferLine += dt.Rows[i][j].ToString();
                    }
                    strmWriterObj.WriteLine(strBufferLine);
                }
                strmWriterObj.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }


        /// <summary>
        /// List转DataTable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="collection"></param>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(IEnumerable<T> collection)
        {
            var props = typeof(T).GetProperties();
            var dt = new DataTable();
            dt.Columns.AddRange(props.Select(p => new DataColumn(p.Name, p.PropertyType)).ToArray());
            if (collection.Count() > 0)
            {
                for (int i = 0; i < collection.Count(); i++)
                {
                    ArrayList tempList = new ArrayList();
                    foreach (PropertyInfo pi in props)
                    {
                        object obj = pi.GetValue(collection.ElementAt(i), null);
                        tempList.Add(obj);
                    }
                    object[] array = tempList.ToArray();
                    dt.LoadDataRow(array, true);
                }
            }
            return dt;
        }



        //接收到串口数据后的解析方法
        public void serpor()
        {
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info+=(str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            
            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(38, 248);//截取str1的1前两个字符

            string input = str2;
            string[] output = Enumerable.Range(0, input.Length / 8)
            .Select(i => input.Substring(i * 8, 8))
            .ToArray();


            #region   解析报文
            //**********把uint换成long类型**************
            //0
            //这一步是把后面四个字符添加到前面
            string inp = output[0];         //这是第一组的8个字符数据
            string oup = inp.Substring(inp.Length - 4) + inp.Substring(0, inp.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins = oup;
            long hex = long.Parse(ins, System.Globalization.NumberStyles.HexNumber);
            float ous = BitConverter.ToSingle(BitConverter.GetBytes(hex), 0);
            //只保留三位小数
            string formattedNum = ous.ToString("F3"); // 保留3位小数并进行四舍五入

            //1
            string inp1 = output[1];         //这是第一组的8个字符数据
            string oup1 = inp1.Substring(inp1.Length - 4) + inp1.Substring(0, inp1.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins1 = oup1;
            long hex1 = long.Parse(ins1, System.Globalization.NumberStyles.HexNumber);
            float ous1 = BitConverter.ToSingle(BitConverter.GetBytes(hex1), 0);
            //只保留三位小数
            string formattedNum1 = ous1.ToString("F3"); // 保留3位小数并进行四舍五入

            //2
            string inp2 = output[2];         //这是第一组的8个字符数据
            string oup2 = inp2.Substring(inp2.Length - 4) + inp2.Substring(0, inp2.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins2 = oup2;
            long hex2 = long.Parse(ins2, System.Globalization.NumberStyles.HexNumber);
            float ous2 = BitConverter.ToSingle(BitConverter.GetBytes(hex2), 0);
            //只保留三位小数
            string formattedNum2 = ous2.ToString("F3"); // 保留3位小数并进行四舍五入


            //3
            string inp3 = output[3];         //这是第一组的8个字符数据
            string oup3 = inp3.Substring(inp3.Length - 4) + inp3.Substring(0, inp3.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins3 = oup3;
            long hex3 = long.Parse(ins3, System.Globalization.NumberStyles.HexNumber);
            float ous3 = BitConverter.ToSingle(BitConverter.GetBytes(hex3), 0);
            //只保留三位小数
            string formattedNum3 = ous3.ToString("F3"); // 保留3位小数并进行四舍五入


            //4
            string inp4 = output[4];         //这是第一组的8个字符数据
            string oup4 = inp4.Substring(inp4.Length - 4) + inp4.Substring(0, inp4.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins4 = oup4;
            long hex4 = long.Parse(ins4, System.Globalization.NumberStyles.HexNumber);
            float ous4 = BitConverter.ToSingle(BitConverter.GetBytes(hex4), 0);
            //只保留三位小数
            string formattedNum4 = ous4.ToString("F3"); // 保留3位小数并进行四舍五入


            //5
            string inp5 = output[5];         //这是第一组的8个字符数据
            string oup5 = inp5.Substring(inp5.Length - 4) + inp5.Substring(0, inp5.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins5 = oup5;
            long hex5 = long.Parse(ins5, System.Globalization.NumberStyles.HexNumber);
            float ous5 = BitConverter.ToSingle(BitConverter.GetBytes(hex5), 0);
            //只保留三位小数
            string formattedNum5 = ous5.ToString("F3"); // 保留3位小数并进行四舍五入



            //6
            string inp6 = output[6];         //这是第一组的8个字符数据
            string oup6 = inp6.Substring(inp6.Length - 5) + inp6.Substring(0, inp6.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins6 = oup6;
            long hex6 = long.Parse(ins6, System.Globalization.NumberStyles.HexNumber);
            float ous6 = BitConverter.ToSingle(BitConverter.GetBytes(hex6), 0);
            //只保留三位小数
            string formattedNum6 = ous6.ToString("F3"); // 保留3位小数并进行四舍五入


            //7
            string inp7 = output[7];         //这是第一组的8个字符数据
            string oup7 = inp7.Substring(inp7.Length - 5) + inp7.Substring(0, inp7.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins7 = oup7;
            long hex7 = long.Parse(ins7, System.Globalization.NumberStyles.HexNumber);
            float ous7 = BitConverter.ToSingle(BitConverter.GetBytes(hex7), 0);
            //只保留三位小数
            string formattedNum7 = ous7.ToString("F3"); // 保留3位小数并进行四舍五入


            //8
            string inp8 = output[8];         //这是第一组的8个字符数据
            string oup8 = inp8.Substring(inp8.Length - 5) + inp8.Substring(0, inp8.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins8 = oup8;
            long hex8 = long.Parse(ins8, System.Globalization.NumberStyles.HexNumber);
            float ous8 = BitConverter.ToSingle(BitConverter.GetBytes(hex8), 0);
            //只保留三位小数
            string formattedNum8 = ous8.ToString("F3"); // 保留3位小数并进行四舍五入


            //9
            string inp9 = output[9];         //这是第一组的8个字符数据
            string oup9 = inp9.Substring(inp9.Length - 5) + inp9.Substring(0, inp9.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins9 = oup9;
            long hex9 = long.Parse(ins9, System.Globalization.NumberStyles.HexNumber);
            float ous9 = BitConverter.ToSingle(BitConverter.GetBytes(hex9), 0);
            //只保留三位小数
            string formattedNum9 = ous9.ToString("F3"); // 保留3位小数并进行四舍五入


            //10
            string inp10 = output[10];         //这是第一组的8个字符数据
            string oup10 = inp10.Substring(inp10.Length - 5) + inp10.Substring(0, inp10.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins10 = oup10;
            long hex10 = long.Parse(ins10, System.Globalization.NumberStyles.HexNumber);
            float ous10 = BitConverter.ToSingle(BitConverter.GetBytes(hex10), 0);
            //只保留三位小数
            string formattedNum10 = ous10.ToString("F3"); // 保留3位小数并进行四舍五入


            //11
            string inp11 = output[11];         //这是第一组的8个字符数据
            string oup11 = inp11.Substring(inp11.Length - 5) + inp11.Substring(0, inp11.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins11 = oup11;
            long hex11 = long.Parse(ins11, System.Globalization.NumberStyles.HexNumber);
            float ous11 = BitConverter.ToSingle(BitConverter.GetBytes(hex11), 0);
            //只保留三位小数
            string formattedNum11 = ous11.ToString("F3"); // 保留3位小数并进行四舍五入


            //12
            string inp12 = output[12];         //这是第一组的8个字符数据
            string oup12 = inp12.Substring(inp12.Length - 5) + inp12.Substring(0, inp12.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins12 = oup12;
            long hex12 = long.Parse(ins12, System.Globalization.NumberStyles.HexNumber);
            float ous12 = BitConverter.ToSingle(BitConverter.GetBytes(hex12), 0);
            //只保留三位小数
            string formattedNum12 = ous12.ToString("F3"); // 保留3位小数并进行四舍五入


            //13
            string inp13 = output[13];         //这是第一组的8个字符数据
            string oup13 = inp13.Substring(inp13.Length - 5) + inp13.Substring(0, inp13.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins13 = oup13;
            long hex13 = long.Parse(ins13, System.Globalization.NumberStyles.HexNumber);
            float ous13 = BitConverter.ToSingle(BitConverter.GetBytes(hex13), 0);
            //只保留三位小数
            string formattedNum13 = ous13.ToString("F3"); // 保留3位小数并进行四舍五入


            //14
            string inp14 = output[14];         //这是第一组的8个字符数据
            string oup14 = inp14.Substring(inp14.Length - 5) + inp14.Substring(0, inp14.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins14 = oup14;
            long hex14 = long.Parse(ins14, System.Globalization.NumberStyles.HexNumber);
            float ous14 = BitConverter.ToSingle(BitConverter.GetBytes(hex14), 0);
            //只保留三位小数
            string formattedNum14 = ous14.ToString("F3"); // 保留3位小数并进行四舍五入


            //15
            string inp15 = output[15];         //这是第一组的8个字符数据
            string oup15 = inp15.Substring(inp15.Length - 5) + inp15.Substring(0, inp15.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins15 = oup15;
            long hex15 = long.Parse(ins15, System.Globalization.NumberStyles.HexNumber);
            float ous15 = BitConverter.ToSingle(BitConverter.GetBytes(hex15), 0);
            //只保留三位小数
            string formattedNum15 = ous15.ToString("F3"); // 保留3位小数并进行四舍五入


            //16
            string inp16 = output[16];         //这是第一组的8个字符数据
            string oup16 = inp16.Substring(inp16.Length - 5) + inp16.Substring(0, inp16.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins16 = oup16;
            long hex16 = long.Parse(ins16, System.Globalization.NumberStyles.HexNumber);
            float ous16 = BitConverter.ToSingle(BitConverter.GetBytes(hex16), 0);
            //只保留三位小数
            string formattedNum16 = ous16.ToString("F3"); // 保留3位小数并进行四舍五入


            //17
            string inp17 = output[17];         //这是第一组的8个字符数据
            string oup17 = inp17.Substring(inp17.Length - 5) + inp17.Substring(0, inp17.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins17 = oup17;
            long hex17 = long.Parse(ins17, System.Globalization.NumberStyles.HexNumber);
            float ous17 = BitConverter.ToSingle(BitConverter.GetBytes(hex17), 0);
            //只保留三位小数
            string formattedNum17 = ous17.ToString("F3"); // 保留3位小数并进行四舍五入



            //18
            string inp18 = output[18];         //这是第一组的8个字符数据
            string oup18 = inp18.Substring(inp18.Length - 5) + inp18.Substring(0, inp18.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins18 = oup18;
            long hex18 = long.Parse(ins18, System.Globalization.NumberStyles.HexNumber);
            float ous18 = BitConverter.ToSingle(BitConverter.GetBytes(hex18), 0);
            //只保留三位小数
            string formattedNum18 = ous18.ToString("F3"); // 保留3位小数并进行四舍五入


            //19
            string inp19 = output[19];         //这是第一组的8个字符数据
            string oup19 = inp19.Substring(inp19.Length - 5) + inp19.Substring(0, inp19.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins19 = oup19;
            long hex19 = long.Parse(ins19, System.Globalization.NumberStyles.HexNumber);
            float ous19 = BitConverter.ToSingle(BitConverter.GetBytes(hex19), 0);
            //只保留三位小数
            string formattedNum19 = ous19.ToString("F3"); // 保留3位小数并进行四舍五入


            //20
            string inp20 = output[20];         //这是第一组的8个字符数据
            string oup20 = inp20.Substring(inp20.Length - 5) + inp20.Substring(0, inp20.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins20 = oup20;
            long hex20 = long.Parse(ins20, System.Globalization.NumberStyles.HexNumber);
            float ous20 = BitConverter.ToSingle(BitConverter.GetBytes(hex20), 0);
            //只保留三位小数
            string formattedNum20 = ous20.ToString("F3"); // 保留3位小数并进行四舍五入


            //21
            string inp21 = output[21];         //这是第一组的8个字符数据
            string oup21 = inp21.Substring(inp21.Length - 5) + inp21.Substring(0, inp21.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins21 = oup21;
            long hex21 = long.Parse(ins21, System.Globalization.NumberStyles.HexNumber);
            float ous21 = BitConverter.ToSingle(BitConverter.GetBytes(hex21), 0);
            //只保留三位小数
            string formattedNum21 = ous21.ToString("F3"); // 保留3位小数并进行四舍五入


            //22
            string inp22 = output[22];         //这是第一组的8个字符数据
            string oup22 = inp22.Substring(inp22.Length - 5) + inp22.Substring(0, inp22.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins22 = oup22;
            long hex22 = long.Parse(ins22, System.Globalization.NumberStyles.HexNumber);
            float ous22 = BitConverter.ToSingle(BitConverter.GetBytes(hex22), 0);
            //只保留三位小数
            string formattedNum22 = ous22.ToString("F3"); // 保留3位小数并进行四舍五入


            //23
            string inp23 = output[23];         //这是第一组的8个字符数据
            string oup23 = inp23.Substring(inp23.Length - 5) + inp23.Substring(0, inp23.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins23 = oup23;
            long hex23 = long.Parse(ins23, System.Globalization.NumberStyles.HexNumber);
            float ous23 = BitConverter.ToSingle(BitConverter.GetBytes(hex23), 0);
            //只保留三位小数
            string formattedNum23 = ous23.ToString("F3"); // 保留3位小数并进行四舍五入


            //24
            string inp24 = output[24];         //这是第一组的8个字符数据
            string oup24 = inp24.Substring(inp24.Length - 5) + inp24.Substring(0, inp24.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins24 = oup24;
            long hex24 = long.Parse(ins24, System.Globalization.NumberStyles.HexNumber);
            float ous24 = BitConverter.ToSingle(BitConverter.GetBytes(hex24), 0);
            //只保留三位小数
            string formattedNum24 = ous24.ToString("F3"); // 保留3位小数并进行四舍五入



            //25
            string inp25 = output[25];         //这是第一组的8个字符数据
            string oup25 = inp25.Substring(inp25.Length - 5) + inp25.Substring(0, inp25.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins25 = oup25;
            long hex25 = long.Parse(ins25, System.Globalization.NumberStyles.HexNumber);
            float ous25 = BitConverter.ToSingle(BitConverter.GetBytes(hex25), 0);
            //只保留三位小数
            string formattedNum25 = ous25.ToString("F3"); // 保留3位小数并进行四舍五入



            //26
            string inp26 = output[26];         //这是第一组的8个字符数据
            string oup26 = inp26.Substring(inp26.Length - 5) + inp26.Substring(0, inp26.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins26 = oup26;
            long hex26 = long.Parse(ins26, System.Globalization.NumberStyles.HexNumber);
            float ous26 = BitConverter.ToSingle(BitConverter.GetBytes(hex26), 0);
            //只保留三位小数
            string formattedNum26 = ous26.ToString("F3"); // 保留3位小数并进行四舍五入


            //27
            string inp27 = output[27];         //这是第一组的8个字符数据
            string oup27 = inp27.Substring(inp27.Length - 5) + inp27.Substring(0, inp27.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins27 = oup27;
            long hex27 = long.Parse(ins27, System.Globalization.NumberStyles.HexNumber);
            float ous27 = BitConverter.ToSingle(BitConverter.GetBytes(hex27), 0);
            //只保留三位小数
            string formattedNum27 = ous27.ToString("F3"); // 保留3位小数并进行四舍五入


            //28
            string inp28 = output[28];         //这是第一组的8个字符数据
            string oup28 = inp28.Substring(inp28.Length - 5) + inp28.Substring(0, inp28.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins28 = oup28;
            long hex28 = long.Parse(ins28, System.Globalization.NumberStyles.HexNumber);
            float ous28 = BitConverter.ToSingle(BitConverter.GetBytes(hex28), 0);
            //只保留三位小数
            string formattedNum28 = ous28.ToString("F3"); // 保留3位小数并进行四舍五入


            // 29
            string inp29 = output[29];         //这是第一组的8个字符数据
            string oup29 = inp29.Substring(inp29.Length - 5) + inp29.Substring(0, inp29.Length - 4);
            //这一步是把16进制转为浮点数格式
            string ins29 = oup29;
            long hex29 = long.Parse(ins29, System.Globalization.NumberStyles.HexNumber);
            float ous29 = BitConverter.ToSingle(BitConverter.GetBytes(hex29), 0);
            //只保留三位小数
            string formattedNum29 = ous29.ToString("F3"); // 保留3位小数并进行四舍五入

            #endregion

            //formattedNum是最后得到的数据
            info = formattedNum.ToString() + "|" + formattedNum1.ToString() + "|" + formattedNum2.ToString()
                + "|" + formattedNum3.ToString() + "|" + formattedNum4.ToString() + "|" + formattedNum5.ToString()
                + "|" + formattedNum6.ToString() + "|" + formattedNum7.ToString() + "|" + formattedNum8.ToString()
                + "|" + formattedNum9.ToString() + "|" + formattedNum10.ToString() + "|" + formattedNum11.ToString()
                + "|" + formattedNum12.ToString() + "|" + formattedNum13.ToString() + "|" + formattedNum14.ToString()
                + "|" + formattedNum15.ToString() + "|" + formattedNum16.ToString() + "|" + formattedNum17.ToString()
            + "|" + formattedNum18.ToString() + "|" + formattedNum19.ToString() + "|" + formattedNum20.ToString()
                + "|" + formattedNum21.ToString() + "|" + formattedNum22.ToString() + "|" + formattedNum23.ToString()
                + "|" + formattedNum24.ToString() + "|" + formattedNum25.ToString() + "|" + formattedNum26.ToString()
                + "|" + formattedNum27.ToString() + "|" + formattedNum28.ToString() + "|" + formattedNum29.ToString();

            #region 变量赋值
            //赋值给变量
            Thread.Sleep(3000);
            //ifs.Zaddress = this.textBox20.Text;
            ifs.Zdianya = float.Parse(formattedNum);
            ifs.Zwendu = float.Parse(formattedNum1);
            ifs.Zallyelvsu = float.Parse(formattedNum2);
            ifs.Zlanzao = float.Parse(formattedNum3);
            ifs.Zlvzao = float.Parse(formattedNum4);
            ifs.Zguizao = float.Parse(formattedNum5);
            ifs.Zjiazao = float.Parse(formattedNum6);
            ifs.Zyinzao = float.Parse(formattedNum7);
            ifs.Zcdom = float.Parse(formattedNum8);
            ifs.Zzhuodu = float.Parse(formattedNum9);
            ifs.Zallswl = float.Parse(formattedNum10);
            ifs.Zlanzaoswl = float.Parse(formattedNum11);
            ifs.Zlvzaoswl = float.Parse(formattedNum12);
            ifs.Zguizaoswl = float.Parse(formattedNum13);
            ifs.Zjiazaoswl = float.Parse(formattedNum14);
            ifs.Zyinzaoswl = float.Parse(formattedNum15);
            ifs.ZF0 = float.Parse(formattedNum20);
            ifs.ZFm = float.Parse(formattedNum21);
            ifs.ZFv = float.Parse(formattedNum22);
            ifs.ZFvFm = float.Parse(formattedNum23);
            ifs.Zsigma = float.Parse(formattedNum24);
            ifs.Zcn = float.Parse(formattedNum26);
            #endregion
            //MessageBox.Show(ifs.Zlanzao.ToString());
            string ddyytt = today.ToString("yyyy-MM-dd");
            //MessageBox.Show(ddyytt);
            Thread.Sleep(1000);
        }

        #region 串口通信检测数据过程

        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                if (textBox20.Text!="" && pictureBox1.Image==null)
                {
                    
                    serpor();
                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox20.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";
                    
                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);
                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();
                    conn.Close();
                    //MessageBox.Show("1号数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox1.Image = Resources.完成;
                    //textBox20.Text = "";
                    Thread.Sleep(3000);

                    if (textBox23.Text != "" || textBox25.Text != "" || textBox27.Text != "" || textBox29.Text != ""
                    || textBox21.Text != "" || textBox22.Text != "" || textBox24.Text != "" || textBox26.Text != ""
                    || textBox28.Text != "" || textBox39.Text != "" || textBox37.Text != "" || textBox35.Text != ""
                    || textBox33.Text != "" || textBox31.Text != "" || textBox38.Text != "" || textBox36.Text != ""
                    || textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
                    {
                        blenderoder2();
                        
                        Thread.Sleep(1000);
                        starttest();
                    }
                }
                

                if (textBox23.Text != "" && pictureBox2.Image == null)
                {
                    serpor();
                    //serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox23.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    //MessageBox.Show("2号数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox2.Image = Resources.完成;
                    Thread.Sleep(3000);
                    if (textBox25.Text != "" || textBox27.Text != "" || textBox29.Text != ""
                    || textBox21.Text != "" || textBox22.Text != "" || textBox24.Text != "" || textBox26.Text != ""
                    || textBox28.Text != "" || textBox39.Text != "" || textBox37.Text != "" || textBox35.Text != ""
                    || textBox33.Text != "" || textBox31.Text != "" || textBox38.Text != "" || textBox36.Text != ""
                    || textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
                    {
                        Thread myThread = new Thread(new ThreadStart(blenderoder3));
                        // 启动线程
                        myThread.Start();
                        Thread.Sleep(1000);
                        myThread.Abort();
                        starttest();
                    }
                }
                

                if (textBox25.Text != "" && pictureBox3.Image == null)
                {
                    serpor();
                    //Thread.Sleep(5000);
                    //serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox25.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    //MessageBox.Show("3号数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox3.Image = Resources.完成;
                    Thread.Sleep(3000);
                    if (textBox27.Text != "" || textBox29.Text != ""
                    || textBox21.Text != "" || textBox22.Text != "" || textBox24.Text != "" || textBox26.Text != ""
                    || textBox28.Text != "" || textBox39.Text != "" || textBox37.Text != "" || textBox35.Text != ""
                    || textBox33.Text != "" || textBox31.Text != "" || textBox38.Text != "" || textBox36.Text != ""
                    || textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
                    {
                        Thread myThread = new Thread(new ThreadStart(blenderoder4));
                        // 启动线程
                        myThread.Start();
                        Thread.Sleep(1000);
                        myThread.Abort();
                        starttest();
                    }
                }


                if (textBox27.Text != "" && pictureBox4.Image == null)
                {
                    
                    serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox27.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    //MessageBox.Show("数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox4.Image = Resources.完成;
                    Thread.Sleep(3000);
                    if (textBox29.Text != ""
                    || textBox21.Text != "" || textBox22.Text != "" || textBox24.Text != "" || textBox26.Text != ""
                    || textBox28.Text != "" || textBox39.Text != "" || textBox37.Text != "" || textBox35.Text != ""
                    || textBox33.Text != "" || textBox31.Text != "" || textBox38.Text != "" || textBox36.Text != ""
                    || textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
                    {
                        Thread myThread = new Thread(new ThreadStart(blenderoder5));
                        // 启动线程
                        myThread.Start();
                        Thread.Sleep(1000);
                        myThread.Abort();
                        starttest();
                    }
                }

                if (textBox29.Text != "" && pictureBox5.Image == null)
                {
                    
                    serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox29.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    //MessageBox.Show("数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox5.Image = Resources.完成;
                    Thread.Sleep(3000);
                    if (textBox21.Text != "" || textBox22.Text != "" || textBox24.Text != "" || textBox26.Text != ""
                    || textBox28.Text != "" || textBox39.Text != "" || textBox37.Text != "" || textBox35.Text != ""
                    || textBox33.Text != "" || textBox31.Text != "" || textBox38.Text != "" || textBox36.Text != ""
                    || textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
                    {
                        Thread myThread = new Thread(new ThreadStart(blenderoder6));
                        // 启动线程
                        myThread.Start();
                        Thread.Sleep(1000);
                        myThread.Abort();
                        starttest();
                    }
                }

                if (textBox21.Text != "" && pictureBox6.Image == null)
                {
                    
                    serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox21.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    //MessageBox.Show("数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox6.Image = Resources.完成;
                    Thread.Sleep(3000);
                    if (textBox22.Text != "" || textBox24.Text != "" || textBox26.Text != ""
                     || textBox28.Text != "" || textBox39.Text != "" || textBox37.Text != "" || textBox35.Text != ""
                     || textBox33.Text != "" || textBox31.Text != "" || textBox38.Text != "" || textBox36.Text != ""
                     || textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
                    {
                        Thread myThread = new Thread(new ThreadStart(blenderoder7));
                        // 启动线程
                        myThread.Start();
                        Thread.Sleep(1000);
                        myThread.Abort();
                        starttest();
                    }
                }

                if (textBox22.Text != "" && pictureBox7.Image == null)
                {
                    
                    serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox22.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    //MessageBox.Show("数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox7.Image = Resources.完成;
                    Thread.Sleep(3000);
                    if (textBox24.Text != "" || textBox26.Text != ""
                     || textBox28.Text != "" || textBox39.Text != "" || textBox37.Text != "" || textBox35.Text != ""
                     || textBox33.Text != "" || textBox31.Text != "" || textBox38.Text != "" || textBox36.Text != ""
                     || textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
                    {
                        Thread myThread = new Thread(new ThreadStart(blenderoder8));
                        // 启动线程
                        myThread.Start();
                        Thread.Sleep(1000);
                        myThread.Abort();
                        starttest();
                    }
                }

                if (textBox24.Text != "" && pictureBox8.Image == null)
                {
                    
                    serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox24.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    // MessageBox.Show("数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox8.Image = Resources.完成;
                    Thread.Sleep(3000);
                    if (textBox26.Text != ""
                     || textBox28.Text != "" || textBox39.Text != "" || textBox37.Text != "" || textBox35.Text != ""
                     || textBox33.Text != "" || textBox31.Text != "" || textBox38.Text != "" || textBox36.Text != ""
                     || textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
                    {
                        Thread myThread = new Thread(new ThreadStart(blenderoder9));
                        // 启动线程
                        myThread.Start();
                        Thread.Sleep(1000);
                        myThread.Abort();
                        starttest();
                    }
                }

                if (textBox26.Text != "" && pictureBox9.Image == null)
                {
                    serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox26.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    // MessageBox.Show("数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox9.Image = Resources.完成;
                    Thread.Sleep(3000);
                    if (textBox28.Text != "" || textBox39.Text != "" || textBox37.Text != "" || textBox35.Text != ""
                     || textBox33.Text != "" || textBox31.Text != "" || textBox38.Text != "" || textBox36.Text != ""
                     || textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
                    {
                        Thread myThread = new Thread(new ThreadStart(blenderoder10));
                        // 启动线程
                        myThread.Start();
                        Thread.Sleep(1000);
                        myThread.Abort();
                        starttest();
                    }
                }

                if (textBox28.Text != "" && pictureBox10.Image == null)
                {
                    
                    serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox28.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    //MessageBox.Show("数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox10.Image = Resources.完成;
                    Thread.Sleep(3000);
                    if (textBox39.Text != "" || textBox37.Text != "" || textBox35.Text != ""
                     || textBox33.Text != "" || textBox31.Text != "" || textBox38.Text != "" || textBox36.Text != ""
                     || textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
                    {
                        Thread myThread = new Thread(new ThreadStart(blenderoder11));
                        // 启动线程
                        myThread.Start();
                        Thread.Sleep(1000);
                        myThread.Abort();
                        starttest();
                    }
                }

                if (textBox39.Text != "" && pictureBox11.Image == null)
                {
                    
                    serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox39.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    //MessageBox.Show("数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox11.Image = Resources.完成;
                    Thread.Sleep(3000);
                    if (textBox37.Text != "" || textBox35.Text != ""
                    || textBox33.Text != "" || textBox31.Text != "" || textBox38.Text != "" || textBox36.Text != ""
                    || textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
                    {
                        Thread myThread = new Thread(new ThreadStart(blenderoder12));
                        // 启动线程
                        myThread.Start();
                        Thread.Sleep(1000);
                        myThread.Abort();
                        starttest();
                    }
                }

                if (textBox37.Text != "" && pictureBox12.Image == null)
                {
                    
                    serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox37.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    // MessageBox.Show("数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox12.Image = Resources.完成;
                    Thread.Sleep(3000);
                    if (textBox35.Text != ""
                    || textBox33.Text != "" || textBox31.Text != "" || textBox38.Text != "" || textBox36.Text != ""
                    || textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
                    {
                        Thread myThread = new Thread(new ThreadStart(blenderoder13));
                        // 启动线程
                        myThread.Start();
                        Thread.Sleep(1000);
                        myThread.Abort();
                        starttest();
                    }
                }

                if (textBox35.Text != "" && pictureBox13.Image == null)
                {
                    
                    serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox35.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    //MessageBox.Show("数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox13.Image = Resources.完成;
                    Thread.Sleep(3000);
                    if (textBox33.Text != "" || textBox31.Text != "" || textBox38.Text != "" || textBox36.Text != ""
                    || textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
                    {
                        Thread myThread = new Thread(new ThreadStart(blenderoder14));
                        // 启动线程
                        myThread.Start();
                        Thread.Sleep(1000);
                        myThread.Abort();
                        starttest();
                    }
                }

                if (textBox33.Text != "" && pictureBox14.Image == null)
                {
                    
                    serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox33.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    // MessageBox.Show("数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox14.Image = Resources.完成;
                    Thread.Sleep(3000);
                    if (textBox31.Text != "" || textBox38.Text != "" || textBox36.Text != ""
                    || textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
                    {
                        Thread myThread = new Thread(new ThreadStart(blenderoder15));
                        // 启动线程
                        myThread.Start();
                        Thread.Sleep(1000);
                        myThread.Abort();
                        starttest();
                    }
                }

                if (textBox31.Text != "" && pictureBox15.Image == null)
                {
                    
                    serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox31.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    //MessageBox.Show("数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox15.Image = Resources.完成;
                    Thread.Sleep(3000);
                    if (textBox38.Text != "" || textBox36.Text != ""
                    || textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
                    {
                        Thread myThread = new Thread(new ThreadStart(blenderoder16));
                        // 启动线程
                        myThread.Start();
                        Thread.Sleep(1000);
                        myThread.Abort();
                        starttest();
                    }
                }

                if (textBox38.Text != "" && pictureBox16.Image == null)
                {
                   
                    serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox38.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    // MessageBox.Show("数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox16.Image = Resources.完成;
                    Thread.Sleep(3000);
                    if (textBox36.Text != ""
                     || textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
                    {
                        Thread myThread = new Thread(new ThreadStart(blenderoder17));
                        // 启动线程
                        myThread.Start();
                        Thread.Sleep(1000);
                        myThread.Abort();
                        starttest();
                    }
                }

                if (textBox36.Text != "" && pictureBox17.Image == null)
                {
                    
                    serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox36.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    // MessageBox.Show("数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox17.Image = Resources.完成;
                    Thread.Sleep(3000);
                    if (textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
                    {
                        Thread myThread = new Thread(new ThreadStart(blenderoder18));
                        // 启动线程
                        myThread.Start();
                        Thread.Sleep(1000);
                        myThread.Abort();
                        starttest();
                    }
                }

                if (textBox34.Text != "" && pictureBox18.Image == null)
                {
                    
                    serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox34.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    //MessageBox.Show("数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox18.Image = Resources.完成;
                    Thread.Sleep(3000);
                    if (textBox32.Text != "" || textBox30.Text != "")
                    {
                        Thread myThread = new Thread(new ThreadStart(blenderoder19));
                        // 启动线程
                        myThread.Start();
                        Thread.Sleep(1000);
                        myThread.Abort();
                        starttest();
                    }
                }

                if (textBox32.Text != "" && pictureBox19.Image == null)
                {
                    
                    serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox32.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    //MessageBox.Show("数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox19.Image = Resources.完成;
                    Thread.Sleep(3000);
                    if (textBox30.Text != "")
                    {
                        Thread myThread = new Thread(new ThreadStart(blenderoder20));
                        // 启动线程
                        myThread.Start();
                        Thread.Sleep(1000);
                        myThread.Abort();
                        starttest();
                    }
                }

                if (textBox30.Text != "" && pictureBox20.Image == null)
                {
                    
                    serpor();

                    //把数据存到数据库
                    // 创建INSERT语句
                    string sqls = "insert into ain(addres,dtimer,dianya,wendu,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl,fo,fm,fv,fvfm,sigma,cn) VALUES ('" + textBox30.Text + "','" + today.ToString() + "'," + ifs.Zdianya + "," + ifs.Zwendu + "," + ifs.Zallyelvsu + "," + ifs.Zlanzao + "," + ifs.Zlvzao + "," + ifs.Zguizao + "," + ifs.Zjiazao + "," + ifs.Zyinzao + "," + ifs.Zcdom + "," + ifs.Zzhuodu + "," + ifs.Zallswl + "," + ifs.Zlanzaoswl + "," + ifs.Zlvzaoswl + "," + ifs.Zguizaoswl + "," + ifs.Zjiazaoswl + "," + ifs.Zyinzaoswl + "," + ifs.ZF0 + "," + ifs.ZFm + "," + ifs.ZFv + "," + ifs.ZFvFm + "," + ifs.Zsigma + "," + ifs.Zcn + ")";

                    // 创建MySQL命令对象
                    MySqlCommand comm1 = new MySqlCommand(sqls, conn);

                    // 打开连接，执行命令并关闭连接
                    conn.Open();
                    comm1.ExecuteNonQuery();

                    conn.Close();
                    //MessageBox.Show("数据检测完成！");
                    info = "";
                    //刷新下拉列表地址选项
                    Thread.Sleep(5000);
                    pictureBox20.Image = Resources.完成;
                    Thread.Sleep(3000);
                    
                }

                
                shuaxinxiala();
                shuaxinzhexiantu();
                Thread.Sleep(2000);
                
                MessageBox.Show("所有样品检测已完成！");

                pictureBox1.Image = null; pictureBox2.Image = null; pictureBox3.Image = null; pictureBox4.Image = null; pictureBox5.Image = null;
                pictureBox6.Image = null; pictureBox7.Image = null; pictureBox8.Image = null; pictureBox9.Image = null; pictureBox10.Image = null;
                pictureBox11.Image = null; pictureBox12.Image = null; pictureBox13.Image = null; pictureBox14.Image = null; pictureBox15.Image = null;
                pictureBox16.Image = null; pictureBox17.Image = null; pictureBox18.Image = null; pictureBox19.Image = null; pictureBox20.Image = null;

                textBox20.Text = string.Empty; textBox23.Text = string.Empty; textBox25.Text = string.Empty; textBox27.Text = string.Empty;
                textBox29.Text = string.Empty; textBox21.Text = string.Empty; textBox22.Text = string.Empty; textBox24.Text = string.Empty;
                textBox26.Text = string.Empty; textBox28.Text = string.Empty; textBox39.Text = string.Empty; textBox37.Text = string.Empty;
                textBox35.Text = string.Empty; textBox33.Text = string.Empty; textBox31.Text = string.Empty; textBox38.Text = string.Empty;
                textBox36.Text = string.Empty; textBox34.Text = string.Empty; textBox32.Text = string.Empty; textBox30.Text = string.Empty;
                textEndtrue();
                materialButton1.Enabled = true;
                materialButton4.Enabled = true;
                materialLabel6.Text = "";

            }
            catch (Exception)
            {
                ;
            }
            Thread.Sleep(2000);

        }
        #endregion

        #region 检测串口拔出
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0219)
            {//设备改变
                if (m.WParam.ToInt32() == 0x8004)
                {//usb串口拔出
                    string[] ports = System.IO.Ports.SerialPort.GetPortNames();//重新获取串口
                    comboBox1.Items.Clear();//清除comboBox里面的数据
                    comboBox1.Items.AddRange(ports);//给comboBox1添加数据
                    if (button1.Text == "关闭串口")
                    {//用户打开过串口
                        if (!serialPort1.IsOpen)
                        {//用户打开的串口被关闭:说明热插拔是用户打开的串口
                            button1.Text = "打开串口";
                            serialPort1.Dispose();//释放掉原先的串口资源
                            comboBox1.SelectedIndex = comboBox1.Items.Count > 0 ? 0 : -1;//显示获取的第一个串口号
                        }
                        else
                        {
                            comboBox1.Text = serialPortName;//显示用户打开的那个串口号
                        }
                    }
                    else
                    {//用户没有打开过串口
                        comboBox1.SelectedIndex = comboBox1.Items.Count > 0 ? 0 : -1;//显示获取的第一个串口号
                    }
                }
                else if (m.WParam.ToInt32() == 0x8000)
                {//usb串口连接上
                    string[] ports = System.IO.Ports.SerialPort.GetPortNames();//重新获取串口
                    comboBox1.Items.Clear();
                    comboBox1.Items.AddRange(ports);
                    if (button1.Text == "关闭串口")
                    {//用户打开过一个串口
                        comboBox1.Text = serialPortName;//显示用户打开的那个串口号
                    }
                    else
                    {
                        comboBox1.SelectedIndex = comboBox1.Items.Count > 0 ? 0 : -1;//显示获取的第一个串口号
                    }
                }
            }
            base.WndProc(ref m);
        }
        #endregion


        #region  无
        private void materialLabel3_Click(object sender, EventArgs e)
        {

        }


        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
           
        }

        
        


       


        private void label55_Click(object sender, EventArgs e)
        {

        }

       

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            
        }
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        #endregion

        
        #region 按条数导出按钮 
        private void materialButton2_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Open();

                string selectedAddress = comboBox2.SelectedItem.ToString();

                if (materialRadioButton4.Checked)
                {
                    string query = "select addres,dtimer,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,fo,fm,fv,fvfm,sigma,cn,wendu,dianya,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl from ain WHERE addres LIKE '%" + selectedAddress + "%' limit 10";
                    MySqlCommand cmd = new MySqlCommand(query, conn);
                    MySqlDataReader reader = cmd.ExecuteReader();
                    //创建Excel工作簿和工作表
                    ExcelPackage excel = new ExcelPackage();

                    var worksheet = excel.Workbook.Worksheets.Add("Sheet1");

                    //写入第一行自定义名称
                    //worksheet.Cells["A1"].Value = "取样地点";
                    worksheet.Cells["A1"].Value = "取样地点";
                    worksheet.Cells["B1"].Value = "检测时间";
                    worksheet.Cells["C1"].Value = "总叶绿素";
                    worksheet.Cells["D1"].Value = "蓝藻";
                    worksheet.Cells["E1"].Value = "绿藻";
                    worksheet.Cells["F1"].Value = "硅藻";
                    worksheet.Cells["G1"].Value = "甲藻";
                    worksheet.Cells["H1"].Value = "隐藻";
                    worksheet.Cells["I1"].Value = "CDOM";
                    worksheet.Cells["J1"].Value = "浊度";
                    worksheet.Cells["K1"].Value = "F0";
                    worksheet.Cells["L1"].Value = "Fm";
                    worksheet.Cells["M1"].Value = "Fv";
                    worksheet.Cells["N1"].Value = "Fv/Fm";
                    worksheet.Cells["O1"].Value = "Sigma";
                    worksheet.Cells["P1"].Value = "Cn";
                    worksheet.Cells["Q1"].Value = "温度";
                    worksheet.Cells["R1"].Value = "电压";
                    worksheet.Cells["S1"].Value = "总生物量";
                    worksheet.Cells["T1"].Value = "蓝藻生物量";
                    worksheet.Cells["U1"].Value = "绿藻生物量";
                    worksheet.Cells["V1"].Value = "硅藻生物量";
                    worksheet.Cells["W1"].Value = "甲藻生物量";
                    worksheet.Cells["X1"].Value = "隐藻生物量";

                    //将查询结果写入Excel中
                    int row = 2;
                    while (reader.Read())
                    {
                        worksheet.Cells["A" + row].Value = reader.GetString(0);
                        worksheet.Cells["B" + row].Value = reader.GetString(1);
                        worksheet.Cells["C" + row].Value = reader.GetString(2);
                        worksheet.Cells["D" + row].Value = reader.GetString(3);
                        worksheet.Cells["E" + row].Value = reader.GetString(4);
                        worksheet.Cells["F" + row].Value = reader.GetString(5);
                        worksheet.Cells["G" + row].Value = reader.GetString(6); 
                        worksheet.Cells["H" + row].Value = reader.GetString(7);
                        worksheet.Cells["I" + row].Value = reader.GetString(8);
                        worksheet.Cells["J" + row].Value = reader.GetString(9);
                        worksheet.Cells["K" + row].Value = reader.GetString(10);
                        worksheet.Cells["L" + row].Value = reader.GetString(11);
                        worksheet.Cells["M" + row].Value = reader.GetString(12);
                        worksheet.Cells["N" + row].Value = reader.GetString(13);
                        worksheet.Cells["O" + row].Value = reader.GetString(14);
                        worksheet.Cells["P" + row].Value = reader.GetString(15);
                        worksheet.Cells["Q" + row].Value = reader.GetString(16);
                        worksheet.Cells["R" + row].Value = reader.GetString(17);
                        worksheet.Cells["S" + row].Value = reader.GetString(18);
                        worksheet.Cells["T" + row].Value = reader.GetString(19);
                        worksheet.Cells["U" + row].Value = reader.GetString(20);
                        worksheet.Cells["V" + row].Value = reader.GetString(21);
                        worksheet.Cells["W" + row].Value = reader.GetString(22);
                        worksheet.Cells["x" + row].Value = reader.GetString(23);
                        row++;
                    }
                    //将Excel文件保存到磁盘上
                    /*excel.SaveAs(new FileInfo("D:\\" + @"" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx"));
                    string path = "D:\\" + @"" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                    MessageBox.Show("导出成功,文件位置:" + path);*/
                    // 保存 Excel 文件
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    saveFileDialog1.Title = "Save Excel file";
                    saveFileDialog1.FileName = comboBox2.Text+ DateTime.Now.ToString("yyyyMMddHHmmss")+".xlsx"; // 设置文件名
                    saveFileDialog1.ShowDialog();

                    if (saveFileDialog1.FileName != "")
                    {
                        // 将 Excel 文件保存到所选位置
                        
                        byte[] bin = excel.GetAsByteArray();
                        File.WriteAllBytes(saveFileDialog1.FileName, bin);
                    }
                    //string path = "D:\\" + @"" + DateTime.Now.ToString("yyyyMMddHHmmss")
                }
                else if (materialRadioButton5.Checked)
                {
                    string query = "select addres,dtimer,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,fo,fm,fv,fvfm,sigma,cn,wendu,dianya,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl from ain WHERE addres LIKE '%" + selectedAddress + "%' limit 50";
                    MySqlCommand cmd = new MySqlCommand(query, conn);
                    MySqlDataReader reader = cmd.ExecuteReader();
                    //创建Excel工作簿和工作表
                    ExcelPackage excel = new ExcelPackage();
                    var worksheet = excel.Workbook.Worksheets.Add("Sheet1");

                    //写入第一行自定义名称
                    //worksheet.Cells["A1"].Value = "取样地点";
                    worksheet.Cells["A1"].Value = "取样地点";
                    worksheet.Cells["B1"].Value = "检测时间";
                    worksheet.Cells["C1"].Value = "总叶绿素";
                    worksheet.Cells["D1"].Value = "蓝藻";
                    worksheet.Cells["E1"].Value = "绿藻";
                    worksheet.Cells["F1"].Value = "硅藻";
                    worksheet.Cells["G1"].Value = "甲藻";
                    worksheet.Cells["H1"].Value = "隐藻";
                    worksheet.Cells["I1"].Value = "CDOM";
                    worksheet.Cells["J1"].Value = "浊度";
                    worksheet.Cells["K1"].Value = "F0";
                    worksheet.Cells["L1"].Value = "Fm";
                    worksheet.Cells["M1"].Value = "Fv";
                    worksheet.Cells["N1"].Value = "Fv/Fm";
                    worksheet.Cells["O1"].Value = "Sigma";
                    worksheet.Cells["P1"].Value = "Cn";
                    worksheet.Cells["Q1"].Value = "温度";
                    worksheet.Cells["R1"].Value = "电压";
                    worksheet.Cells["S1"].Value = "总生物量";
                    worksheet.Cells["T1"].Value = "蓝藻生物量";
                    worksheet.Cells["U1"].Value = "绿藻生物量";
                    worksheet.Cells["V1"].Value = "硅藻生物量";
                    worksheet.Cells["W1"].Value = "甲藻生物量";
                    worksheet.Cells["X1"].Value = "隐藻生物量";

                    //将查询结果写入Excel中
                    int row = 2;
                    while (reader.Read())
                    {
                        worksheet.Cells["A" + row].Value = reader.GetString(0);
                        worksheet.Cells["B" + row].Value = reader.GetString(1);
                        worksheet.Cells["C" + row].Value = reader.GetString(2);
                        worksheet.Cells["D" + row].Value = reader.GetString(3);
                        worksheet.Cells["E" + row].Value = reader.GetString(4);
                        worksheet.Cells["F" + row].Value = reader.GetString(5);
                        worksheet.Cells["G" + row].Value = reader.GetString(6);
                        worksheet.Cells["H" + row].Value = reader.GetString(7);
                        worksheet.Cells["I" + row].Value = reader.GetString(8);
                        worksheet.Cells["J" + row].Value = reader.GetString(9);
                        worksheet.Cells["K" + row].Value = reader.GetString(10);
                        worksheet.Cells["L" + row].Value = reader.GetString(11);
                        worksheet.Cells["M" + row].Value = reader.GetString(12);
                        worksheet.Cells["N" + row].Value = reader.GetString(13);
                        worksheet.Cells["O" + row].Value = reader.GetString(14);
                        worksheet.Cells["P" + row].Value = reader.GetString(15);
                        worksheet.Cells["Q" + row].Value = reader.GetString(16);
                        worksheet.Cells["R" + row].Value = reader.GetString(17);
                        worksheet.Cells["S" + row].Value = reader.GetString(18);
                        worksheet.Cells["T" + row].Value = reader.GetString(19);
                        worksheet.Cells["U" + row].Value = reader.GetString(20);
                        worksheet.Cells["V" + row].Value = reader.GetString(21);
                        worksheet.Cells["W" + row].Value = reader.GetString(22);
                        worksheet.Cells["x" + row].Value = reader.GetString(23);
                        row++;
                    }
                    //将Excel文件保存到磁盘上
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    saveFileDialog1.Title = "Save Excel file";
                    saveFileDialog1.FileName = comboBox2.Text  + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx"; // 设置文件名
                    saveFileDialog1.ShowDialog();

                    if (saveFileDialog1.FileName != "")
                    {
                        // 将 Excel 文件保存到所选位置

                        byte[] bin = excel.GetAsByteArray();
                        File.WriteAllBytes(saveFileDialog1.FileName, bin);
                    }
                }
                else if (materialRadioButton6.Checked)
                {
                    string query = "select addres,dtimer,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,fo,fm,fv,fvfm,sigma,cn,wendu,dianya,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl from ain WHERE addres LIKE '%" + selectedAddress + "%' limit 100";
                    MySqlCommand cmd = new MySqlCommand(query, conn);
                    MySqlDataReader reader = cmd.ExecuteReader();
                    //创建Excel工作簿和工作表
                    ExcelPackage excel = new ExcelPackage();
                    var worksheet = excel.Workbook.Worksheets.Add("Sheet1");

                    //写入第一行自定义名称
                    //worksheet.Cells["A1"].Value = "取样地点";
                    worksheet.Cells["A1"].Value = "取样地点";
                    worksheet.Cells["B1"].Value = "检测时间";
                    worksheet.Cells["C1"].Value = "总叶绿素";
                    worksheet.Cells["D1"].Value = "蓝藻";
                    worksheet.Cells["E1"].Value = "绿藻";
                    worksheet.Cells["F1"].Value = "硅藻";
                    worksheet.Cells["G1"].Value = "甲藻";
                    worksheet.Cells["H1"].Value = "隐藻";
                    worksheet.Cells["I1"].Value = "CDOM";
                    worksheet.Cells["J1"].Value = "浊度";
                    worksheet.Cells["K1"].Value = "F0";
                    worksheet.Cells["L1"].Value = "Fm";
                    worksheet.Cells["M1"].Value = "Fv";
                    worksheet.Cells["N1"].Value = "Fv/Fm";
                    worksheet.Cells["O1"].Value = "Sigma";
                    worksheet.Cells["P1"].Value = "Cn";
                    worksheet.Cells["Q1"].Value = "温度";
                    worksheet.Cells["R1"].Value = "电压";
                    worksheet.Cells["S1"].Value = "总生物量";
                    worksheet.Cells["T1"].Value = "蓝藻生物量";
                    worksheet.Cells["U1"].Value = "绿藻生物量";
                    worksheet.Cells["V1"].Value = "硅藻生物量";
                    worksheet.Cells["W1"].Value = "甲藻生物量";
                    worksheet.Cells["X1"].Value = "隐藻生物量";

                    //将查询结果写入Excel中
                    int row = 2;
                    while (reader.Read())
                    {
                        worksheet.Cells["A" + row].Value = reader.GetString(0);
                        worksheet.Cells["B" + row].Value = reader.GetString(1);
                        worksheet.Cells["C" + row].Value = reader.GetString(2);
                        worksheet.Cells["D" + row].Value = reader.GetString(3);
                        worksheet.Cells["E" + row].Value = reader.GetString(4);
                        worksheet.Cells["F" + row].Value = reader.GetString(5);
                        worksheet.Cells["G" + row].Value = reader.GetString(6);
                        worksheet.Cells["H" + row].Value = reader.GetString(7);
                        worksheet.Cells["I" + row].Value = reader.GetString(8);
                        worksheet.Cells["J" + row].Value = reader.GetString(9);
                        worksheet.Cells["K" + row].Value = reader.GetString(10);
                        worksheet.Cells["L" + row].Value = reader.GetString(11);
                        worksheet.Cells["M" + row].Value = reader.GetString(12);
                        worksheet.Cells["N" + row].Value = reader.GetString(13);
                        worksheet.Cells["O" + row].Value = reader.GetString(14);
                        worksheet.Cells["P" + row].Value = reader.GetString(15);
                        worksheet.Cells["Q" + row].Value = reader.GetString(16);
                        worksheet.Cells["R" + row].Value = reader.GetString(17);
                        worksheet.Cells["S" + row].Value = reader.GetString(18);
                        worksheet.Cells["T" + row].Value = reader.GetString(19);
                        worksheet.Cells["U" + row].Value = reader.GetString(20);
                        worksheet.Cells["V" + row].Value = reader.GetString(21);
                        worksheet.Cells["W" + row].Value = reader.GetString(22);
                        worksheet.Cells["x" + row].Value = reader.GetString(23);
                        row++;
                    }
                    //将Excel文件保存到磁盘上
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    saveFileDialog1.Title = "Save Excel file";
                    saveFileDialog1.FileName = comboBox2.Text  + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx"; // 设置文件名
                    saveFileDialog1.ShowDialog();

                    if (saveFileDialog1.FileName != "")
                    {
                        // 将 Excel 文件保存到所选位置

                        byte[] bin = excel.GetAsByteArray();
                        File.WriteAllBytes(saveFileDialog1.FileName, bin);
                    }
                }
                else
                {
                    MessageBox.Show("请选择需要导出的数据条数");
                }

                //关闭MySQL连接
                conn.Close();
            }
            catch (Exception)
            {

                MessageBox.Show("请选择地址!");
            }
            
        }
        #endregion

        

        #region 导出当前
        private void materialButton3_Click(object sender, EventArgs e)
        {
            if (countts>0)
            {
                conn.Open();
                string query = "select addres,dtimer,allyls,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,fo,fm,fv,fvfm,sigma,cn,wendu,dianya,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl from ain ORDER BY dtimer DESC LIMIT " + countts + "";
                MySqlCommand cmd = new MySqlCommand(query, conn);
                MySqlDataReader reader = cmd.ExecuteReader();
                //创建Excel工作簿和工作表
                ExcelPackage excel = new ExcelPackage();

                var worksheet = excel.Workbook.Worksheets.Add("Sheet1");

                //写入第一行自定义名称
                //worksheet.Cells["A1"].Value = "取样地点";
                worksheet.Cells["A1"].Value = "取样地点";
                worksheet.Cells["B1"].Value = "检测时间";
                worksheet.Cells["C1"].Value = "总叶绿素";
                worksheet.Cells["D1"].Value = "蓝藻";
                worksheet.Cells["E1"].Value = "绿藻";
                worksheet.Cells["F1"].Value = "硅藻";
                worksheet.Cells["G1"].Value = "甲藻";
                worksheet.Cells["H1"].Value = "隐藻";
                worksheet.Cells["I1"].Value = "CDOM";
                worksheet.Cells["J1"].Value = "浊度";
                worksheet.Cells["K1"].Value = "F0";
                worksheet.Cells["L1"].Value = "Fm";
                worksheet.Cells["M1"].Value = "Fv";
                worksheet.Cells["N1"].Value = "Fv/Fm";
                worksheet.Cells["O1"].Value = "Sigma";
                worksheet.Cells["P1"].Value = "Cn";
                worksheet.Cells["Q1"].Value = "温度";
                worksheet.Cells["R1"].Value = "电压";
                worksheet.Cells["S1"].Value = "总生物量";
                worksheet.Cells["T1"].Value = "蓝藻生物量";
                worksheet.Cells["U1"].Value = "绿藻生物量";
                worksheet.Cells["V1"].Value = "硅藻生物量";
                worksheet.Cells["W1"].Value = "甲藻生物量";
                worksheet.Cells["X1"].Value = "隐藻生物量";

                //将查询结果写入Excel中
                int row = 2;
                while (reader.Read())
                {
                    worksheet.Cells["A" + row].Value = reader.GetString(0);
                    worksheet.Cells["B" + row].Value = reader.GetString(1);
                    worksheet.Cells["C" + row].Value = reader.GetString(2);
                    worksheet.Cells["D" + row].Value = reader.GetString(3);
                    worksheet.Cells["E" + row].Value = reader.GetString(4);
                    worksheet.Cells["F" + row].Value = reader.GetString(5);
                    worksheet.Cells["G" + row].Value = reader.GetString(6);
                    worksheet.Cells["H" + row].Value = reader.GetString(7);
                    worksheet.Cells["I" + row].Value = reader.GetString(8);
                    worksheet.Cells["J" + row].Value = reader.GetString(9);
                    worksheet.Cells["K" + row].Value = reader.GetString(10);
                    worksheet.Cells["L" + row].Value = reader.GetString(11);
                    worksheet.Cells["M" + row].Value = reader.GetString(12);
                    worksheet.Cells["N" + row].Value = reader.GetString(13);
                    worksheet.Cells["O" + row].Value = reader.GetString(14);
                    worksheet.Cells["P" + row].Value = reader.GetString(15);
                    worksheet.Cells["Q" + row].Value = reader.GetString(16);
                    worksheet.Cells["R" + row].Value = reader.GetString(17);
                    worksheet.Cells["S" + row].Value = reader.GetString(18);
                    worksheet.Cells["T" + row].Value = reader.GetString(19);
                    worksheet.Cells["U" + row].Value = reader.GetString(20);
                    worksheet.Cells["V" + row].Value = reader.GetString(21);
                    worksheet.Cells["W" + row].Value = reader.GetString(22);
                    worksheet.Cells["x" + row].Value = reader.GetString(23);
                    row++;
                }
                //将Excel文件保存到磁盘上
                /*excel.SaveAs(new FileInfo("D:\\" + @"" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx"));
                string path = "D:\\" + @"" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                MessageBox.Show("导出成功,文件位置:" + path);*/
                // 保存 Excel 文件
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveFileDialog1.Title = "Save Excel file";
                //saveFileDialog1.FileName = "当前" + "|" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx"; // 设置文件名
                saveFileDialog1.ShowDialog();

                if (saveFileDialog1.FileName != "")
                {
                    // 将 Excel 文件保存到所选位置

                    byte[] bin = excel.GetAsByteArray();
                    File.WriteAllBytes(saveFileDialog1.FileName, bin);
                }
            }
            else
            {
                MessageBox.Show("还未检测过数据！");
            }

            conn.Close();
            /*DataTable dt = new DataTable();
            dt.Columns.Add("aa");
            dt.Columns.Add("bb");
            dt.Columns.Add("cc");
            dt.Columns.Add("dd");
            dt.Columns.Add("ee");
            dt.Columns.Add("ff");
            dt.Columns.Add("gg");
            dt.Columns.Add("hh");
            dt.Columns.Add("ii");
            dt.Columns.Add("jj");
            dt.Columns.Add("kk");
            dt.Columns.Add("ll");
            dt.Columns.Add("mm");
            dt.Columns.Add("nn");
            dt.Columns.Add("oo");
            dt.Columns.Add("pp");
            dt.Columns.Add("qq");
            dt.Columns.Add("rr");
            dt.Columns.Add("ss");
            dt.Columns.Add("tt");
            dt.Columns.Add("uu");
            dt.Columns.Add("vv");
            dt.Columns.Add("ww");
            dt.Columns.Add("xx");
            dt.Columns.Add("yy");

            //这里给各个测试数据赋值
            DataRow dr = dt.NewRow();
            dr[0] = comboBox2.Text;
            dr[1] = label1.Text;
            dr[2] = label48.Text;
            //dr[3] = textBox4.Text;
            dr[3] = textBox1.Text;
            dr[4] = textBox2.Text;
            dr[5] = textBox3.Text;
            dr[6] = textBox4.Text;
            dr[7] = textBox5.Text;
            dr[8] = textBox18.Text;
            dr[9] = textBox19.Text;
            dr[10] = textBox15.Text;
            dr[11] = textBox14.Text;
            dr[12] = textBox13.Text;
            dr[13] = label22.Text;
            dr[14] = textBox12.Text;
            dr[15] = textBox11.Text;
            dr[16] = textBox16.Text;
            dr[17] = textBox17.Text;
            dr[18] = label14.Text;
            dr[19] = textBox10.Text;
            dr[20] = textBox9.Text;
            dr[21] = textBox8.Text;
            dr[22] = textBox7.Text;
            dr[23] = textBox6.Text;

            dt.Rows.Add(dr);
            //这里是添加测试数据的名称
            string path = "D:\\" + @"" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            if (dt2csv(dt, path, "藻类信息", "取样地点,时间,总叶绿素,蓝藻,绿藻,硅藻,甲藻,隐藻,CDOM,浊度,F0,Fm,Fv,Fv/Fm,Sigma,Cn,温度,电压,总生物量,蓝藻生物量,绿藻生物量,硅藻生物量,甲藻生物量,隐藻生物量,"))
            {
                MessageBox.Show("导出成功,文件位置:" + path);
            }
            else
            {
                MessageBox.Show("导出失败");
            }*/
        }
        #endregion


        public void xxinfoseripot()
        {
            try
            {//防止意外错误
                serialPort1.PortName = comboBox1.Text;//获取comboBox1要打开的串口号
                serialPortName = comboBox1.Text;
                serialPort1.BaudRate = int.Parse(comboBox3.Text);//获取comboBox2选择的波特率
                serialPort1.DataBits = int.Parse(comboBox5.Text);//设置数据位
                /*设置停止位*/
                if (comboBox4.Text == "1") { serialPort1.StopBits = StopBits.One; }
                else if (comboBox4.Text == "1.5") { serialPort1.StopBits = StopBits.OnePointFive; }
                else if (comboBox4.Text == "2") { serialPort1.StopBits = StopBits.Two; }
                /*设置奇偶校验*/
                if (comboBox6.Text == "无") { serialPort1.Parity = Parity.None; }
                else if (comboBox6.Text == "奇校验") { serialPort1.Parity = Parity.Odd; }
                else if (comboBox6.Text == "偶校验") { serialPort1.Parity = Parity.Even; }

                serialPort1.Open();//打开串口
                //button1.Text = "关闭串口";//按钮显示关闭串口
                Thread.Sleep(1000);
               

            }
            catch (Exception err)
            {
                MessageBox.Show("未检测到串口，确保正常连接！");//对话框显示打开失败
            }
        }

        public void textEndfalse()
        {
            textBox20.Enabled = false; textBox23.Enabled = false; textBox25.Enabled = false; textBox27.Enabled = false; textBox29.Enabled = false;
            textBox21.Enabled = false; textBox22.Enabled = false; textBox24.Enabled = false; textBox26.Enabled = false; textBox28.Enabled = false;
            textBox39.Enabled = false; textBox37.Enabled = false; textBox35.Enabled = false; textBox33.Enabled = false; textBox31.Enabled = false;
            textBox38.Enabled = false; textBox36.Enabled = false; textBox34.Enabled = false; textBox32.Enabled = false; textBox30.Enabled = false;
        }

        public void textEndtrue()
        {
            textBox20.Enabled = true; textBox23.Enabled = true; textBox25.Enabled = true; textBox27.Enabled = true; textBox29.Enabled = true;
            textBox21.Enabled = true; textBox22.Enabled = true; textBox24.Enabled = true; textBox26.Enabled = true; textBox28.Enabled = true;
            textBox39.Enabled = true; textBox37.Enabled = true; textBox35.Enabled = true; textBox33.Enabled = true; textBox31.Enabled = true;
            textBox38.Enabled = true; textBox36.Enabled = true; textBox34.Enabled = true; textBox32.Enabled = true; textBox30.Enabled = true;
        }

        //开始检测按钮
        private void materialButton4_Click(object sender, EventArgs e)
        {
            if (textBox20.Text != "" || textBox23.Text != "" || textBox25.Text != "" || textBox27.Text != "" || textBox29.Text != ""
                    || textBox21.Text != "" || textBox22.Text != "" || textBox24.Text != "" || textBox26.Text != ""
                    || textBox28.Text != "" || textBox39.Text != "" || textBox37.Text != "" || textBox35.Text != ""
                    || textBox33.Text != "" || textBox31.Text != "" || textBox38.Text != "" || textBox36.Text != ""
                    || textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
            {
                //复位
                //resetxxx();
                //Thread.Sleep(3000);
                //开启电机 阀门
                /*Thread myThread = new Thread(new ThreadStart(blenderoder1));
                // 启动线程
                myThread.Start();*/
                /*Task task = Task.Run(() => {
                    
                });*/

                //task.Wait(); // 等待异步任务完成
                blenderoder1();
                //Thread.Sleep(20000);
                //myThread.Abort();


                int count = 0; // 用于计数的变量

                string[] textBoxNames = { "textBox20", "textBox23", "textBox25", "textBox27", "textBox29", "textBox21",
                          "textBox22", "textBox24", "textBox26", "textBox28", "textBox39", "textBox37",
                          "textBox35", "textBox33", "textBox31", "textBox38", "textBox36", "textBox34",
                          "textBox32", "textBox30" }; // 存储所有文本框名称的数组

                foreach (string textBoxName in textBoxNames) // 遍历所有文本框
                {
                    System.Windows.Forms.TextBox textBox = this.Controls.Find(textBoxName, true).FirstOrDefault() as System.Windows.Forms.TextBox; // 获取文本框控件

                    if (textBox != null && !string.IsNullOrEmpty(textBox.Text)) // 判断文本框不为空
                    {
                        count++; // 增加计数器

                    }
                }
                countts = count;
                
                Thread.Sleep(1000);
                //开始检测
                starttest();
                //禁止输入
                textEndfalse();

                materialButton4.Enabled = false;
                materialButton1.Enabled = false;
            }
            else
            {
                MessageBox.Show("请至少输入一个地址");
            }
        }
        
        //搅拌机命令
        public void blenderoder1()
        {
            Byte[] buffer3 = new Byte[8];
            buffer3[0] = 0xCC;
            buffer3[1] = 0x00;
            buffer3[2] = 0x45;
            buffer3[3] = 0x00;
            buffer3[4] = 0x00;
            buffer3[5] = 0xDD;
            buffer3[6] = 0xEE;
            buffer3[7] = 0x01;
            serialPort1.Write(buffer3, 0, 8);

            Byte[] buffer4 = new Byte[8];
            buffer4[0] = 0xCC;
            buffer4[1] = 0x01;
            buffer4[2] = 0x45;
            buffer4[3] = 0x00;
            buffer4[4] = 0x00;
            buffer4[5] = 0xDD;
            buffer4[6] = 0xEF;
            buffer4[7] = 0x01;
            serialPort1.Write(buffer4, 0, 8);
            Thread.Sleep(3000);
            //电机打开
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x02;
            buffer[2] = 0xA4;
            buffer[3] = 0x01;
            buffer[4] = 0x0A;
            buffer[5] = 0xDD;
            buffer[6] = 0x5A;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(10000);
            //电机关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x02;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6A;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            //1号阀门打开
            Byte[] buffer2 = new Byte[8];
            buffer2[0] = 0xCC;
            buffer2[1] = 0x01;
            buffer2[2] = 0xA4;
            buffer2[3] = 0x01;
            buffer2[4] = 0x0A;
            buffer2[5] = 0xDD;
            buffer2[6] = 0x59;
            buffer2[7] = 0x02;
            serialPort1.Write(buffer2, 0, 8);
            Thread.Sleep(1000);
            

            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }
            
            
            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    materialLabel6.Text = "正常执行"; break;
                case "FF":
                    materialLabel6.Text = "未知错误"; break;
                case "06":
                    materialLabel6.Text = "未知位置"; break;
                case "05":
                    materialLabel6.Text = "电机堵转"; break;
                case "04":
                    materialLabel6.Text = "电机忙"; break;
                case "03":
                    materialLabel6.Text = "光耦错误"; break;
                case "02":
                    materialLabel6.Text = "参数错误"; break;
                case "01":
                    materialLabel6.Text = "桢错误"; break;
                case "00":
                    materialLabel6.Text = "正常"; break;
                default:
                    materialLabel6.Text = "未返回任何命令"; break;

            }
            Thread.Sleep(2000);

            info = "";

        }

        public void blenderoder2()
        {
            //电机打开
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x02;
            buffer[2] = 0xA4;
            buffer[3] = 0x02;
            buffer[4] = 0x01;
            buffer[5] = 0xDD;
            buffer[6] = 0x52;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(10000);
            //电机关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x02;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6A;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);
            //1号阀门打开
            Byte[] buffer2 = new Byte[8];
            buffer2[0] = 0xCC;
            buffer2[1] = 0x00;
            buffer2[2] = 0xA4;
            buffer2[3] = 0x02;
            buffer2[4] = 0x01;
            buffer2[5] = 0xDD;
            buffer2[6] = 0x50;
            buffer2[7] = 0x02;
            serialPort1.Write(buffer2, 0, 8);
            Thread.Sleep(1000);
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";
        }

        public void blenderoder3()
        {
            //电机打开
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x02;
            buffer[2] = 0xA4;
            buffer[3] = 0x03;
            buffer[4] = 0x02;
            buffer[5] = 0xDD;
            buffer[6] = 0x54;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(10000);
            //电机关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x02;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6A;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer3 = new Byte[8];
            buffer3[0] = 0xCC;
            buffer3[1] = 0x00;
            buffer3[2] = 0xA4;
            buffer3[3] = 0x03;
            buffer3[4] = 0x02;
            buffer3[5] = 0xDD;
            buffer3[6] = 0x52;
            buffer3[7] = 0x02;
            serialPort1.Write(buffer3, 0, 8);
            Thread.Sleep(1000);
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";
        }

        public void blenderoder4()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x02;
            buffer[2] = 0xA4;
            buffer[3] = 0x04;
            buffer[4] = 0x03;
            buffer[5] = 0xDD;
            buffer[6] = 0x56;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(1000);
            //关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x02;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6A;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer3 = new Byte[8];
            buffer3[0] = 0xCC;
            buffer3[1] = 0x00;
            buffer3[2] = 0xA4;
            buffer3[3] = 0x04;
            buffer3[4] = 0x03;
            buffer3[5] = 0xDD;
            buffer3[6] = 0x54;
            buffer3[7] = 0x02;
            serialPort1.Write(buffer3, 0, 8);
            Thread.Sleep(1000);
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";
        }
        public void blenderoder5()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x02;
            buffer[2] = 0xA4;
            buffer[3] = 0x05;
            buffer[4] = 0x04;
            buffer[5] = 0xDD;
            buffer[6] = 0x58;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(1000);
            //关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x02;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6A;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer3 = new Byte[8];
            buffer3[0] = 0xCC;
            buffer3[1] = 0x00;
            buffer3[2] = 0xA4;
            buffer3[3] = 0x05;
            buffer3[4] = 0x04;
            buffer3[5] = 0xDD;
            buffer3[6] = 0x56;
            buffer3[7] = 0x02;
            serialPort1.Write(buffer3, 0, 8);
            Thread.Sleep(1000);
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";
        }

        public void blenderoder6()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x02;
            buffer[2] = 0xA4;
            buffer[3] = 0x06;
            buffer[4] = 0x05;
            buffer[5] = 0xDD;
            buffer[6] = 0x5A;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(1000);
            //关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x02;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6A;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer3 = new Byte[8];
            buffer3[0] = 0xCC;
            buffer3[1] = 0x00;
            buffer3[2] = 0xA4;
            buffer3[3] = 0x06;
            buffer3[4] = 0x05;
            buffer3[5] = 0xDD;
            buffer3[6] = 0x58;
            buffer3[7] = 0x02;
            serialPort1.Write(buffer3, 0, 8);
            Thread.Sleep(1000);
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";
        }

        public void blenderoder7()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x02;
            buffer[2] = 0xA4;
            buffer[3] = 0x07;
            buffer[4] = 0x06;
            buffer[5] = 0xDD;
            buffer[6] = 0x5C;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(1000);
            //关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x02;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6A;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer3 = new Byte[8];
            buffer3[0] = 0xCC;
            buffer3[1] = 0x00;
            buffer3[2] = 0xA4;
            buffer3[3] = 0x07;
            buffer3[4] = 0x06;
            buffer3[5] = 0xDD;
            buffer3[6] = 0x5A;
            buffer3[7] = 0x02;
            serialPort1.Write(buffer3, 0, 8);
            Thread.Sleep(1000);
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";
        }

        public void blenderoder8()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x02;
            buffer[2] = 0xA4;
            buffer[3] = 0x08;
            buffer[4] = 0x07;
            buffer[5] = 0xDD;
            buffer[6] = 0x5E;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(1000);
            //关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x02;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6A;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer3 = new Byte[8];
            buffer3[0] = 0xCC;
            buffer3[1] = 0x00;
            buffer3[2] = 0xA4;
            buffer3[3] = 0x08;
            buffer3[4] = 0x07;
            buffer3[5] = 0xDD;
            buffer3[6] = 0x5C;
            buffer3[7] = 0x02;
            serialPort1.Write(buffer3, 0, 8);
            Thread.Sleep(1000);
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";
        }

        public void blenderoder9()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x02;
            buffer[2] = 0xA4;
            buffer[3] = 0x09;
            buffer[4] = 0x08;
            buffer[5] = 0xDD;
            buffer[6] = 0x60;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(1000);
            //关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x02;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6A;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer3 = new Byte[8];
            buffer3[0] = 0xCC;
            buffer3[1] = 0x00;
            buffer3[2] = 0xA4;
            buffer3[3] = 0x09;
            buffer3[4] = 0x08;
            buffer3[5] = 0xDD;
            buffer3[6] = 0x5E;
            buffer3[7] = 0x02;
            serialPort1.Write(buffer3, 0, 8);
            Thread.Sleep(1000);
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";
        }

        public void blenderoder10()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x02;
            buffer[2] = 0xA4;
            buffer[3] = 0x0A;
            buffer[4] = 0x09;
            buffer[5] = 0xDD;
            buffer[6] = 0x62    ;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(1000);
            //关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x02;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6A;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer3 = new Byte[8];
            buffer3[0] = 0xCC;
            buffer3[1] = 0x00;
            buffer3[2] = 0xA4;
            buffer3[3] = 0x0A;
            buffer3[4] = 0x09;
            buffer3[5] = 0xDD;
            buffer3[6] = 0x60;
            buffer3[7] = 0x02;
            serialPort1.Write(buffer3, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer4 = new Byte[8];
            buffer4[0] = 0xCC;
            buffer4[1] = 0x00;
            buffer4[2] = 0xB4;
            buffer4[3] = 0x0A;
            buffer4[4] = 0x01;
            buffer4[5] = 0xDD;
            buffer4[6] = 0x68;
            buffer4[7] = 0x02;
            serialPort1.Write(buffer4, 0, 8);
            Thread.Sleep(1000);
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";
        }

        public void blenderoder11()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x03;
            buffer[2] = 0xA4;
            buffer[3] = 0x01;
            buffer[4] = 0x0A;
            buffer[5] = 0xDD;
            buffer[6] = 0x5B;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(1000);
            //关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x03;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6B;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer2 = new Byte[8];
            buffer2[0] = 0xCC;
            buffer2[1] = 0x01;
            buffer2[2] = 0xA4;
            buffer2[3] = 0x01;
            buffer2[4] = 0x0A;
            buffer2[5] = 0xDD;
            buffer2[6] = 0x59;
            buffer2[7] = 0x02;
            serialPort1.Write(buffer2, 0, 8);
            Thread.Sleep(1000);
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";

        }

        public void blenderoder12()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x03;
            buffer[2] = 0xA4;
            buffer[3] = 0x02;
            buffer[4] = 0x01;
            buffer[5] = 0xDD;
            buffer[6] = 0x53;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(1000);
            //关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x03;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6B;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer2 = new Byte[8];
            buffer2[0] = 0xCC;
            buffer2[1] = 0x01;
            buffer2[2] = 0xA4;
            buffer2[3] = 0x02;
            buffer2[4] = 0x01;
            buffer2[5] = 0xDD;
            buffer2[6] = 0x51;
            buffer2[7] = 0x02;
            serialPort1.Write(buffer2, 0, 8);
            Thread.Sleep(1000);
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";
        }

        public void blenderoder13()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x03;
            buffer[2] = 0xA4;
            buffer[3] = 0x03;
            buffer[4] = 0x02;
            buffer[5] = 0xDD;
            buffer[6] = 0x55;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(1000);
            //关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x03;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6B;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer2 = new Byte[8];
            buffer2[0] = 0xCC;
            buffer2[1] = 0x01;
            buffer2[2] = 0xA4;
            buffer2[3] = 0x03;
            buffer2[4] = 0x02;
            buffer2[5] = 0xDD;
            buffer2[6] = 0x53;
            buffer2[7] = 0x02;
            serialPort1.Write(buffer2, 0, 8);
            Thread.Sleep(1000);
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";
        }

        public void blenderoder14()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x03;
            buffer[2] = 0xA4;
            buffer[3] = 0x04;
            buffer[4] = 0x03;
            buffer[5] = 0xDD;
            buffer[6] = 0x57;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(1000);
            //关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x03;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6B;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer2 = new Byte[8];
            buffer2[0] = 0xCC;
            buffer2[1] = 0x01;
            buffer2[2] = 0xA4;
            buffer2[3] = 0x04;
            buffer2[4] = 0x03;
            buffer2[5] = 0xDD;
            buffer2[6] = 0x55;
            buffer2[7] = 0x02;
            serialPort1.Write(buffer2, 0, 8);
            Thread.Sleep(1000);
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";
        }

        public void blenderoder15()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x03;
            buffer[2] = 0xA4;
            buffer[3] = 0x05;
            buffer[4] = 0x04;
            buffer[5] = 0xDD;
            buffer[6] = 0x59;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(1000);
            //关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x03;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6B;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer2 = new Byte[8];
            buffer2[0] = 0xCC;
            buffer2[1] = 0x01;
            buffer2[2] = 0xA4;
            buffer2[3] = 0x05;
            buffer2[4] = 0x04;
            buffer2[5] = 0xDD;
            buffer2[6] = 0x57;
            buffer2[7] = 0x02;
            serialPort1.Write(buffer2, 0, 8);
            Thread.Sleep(1000);
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";
        }

        public void blenderoder16()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x03;
            buffer[2] = 0xA4;
            buffer[3] = 0x06;
            buffer[4] = 0x05;
            buffer[5] = 0xDD;
            buffer[6] = 0x5B;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(1000);
            //关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x03;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6B;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer2 = new Byte[8];
            buffer2[0] = 0xCC;
            buffer2[1] = 0x01;
            buffer2[2] = 0xA4;
            buffer2[3] = 0x06;
            buffer2[4] = 0x05;
            buffer2[5] = 0xDD;
            buffer2[6] = 0x59;
            buffer2[7] = 0x02;
            serialPort1.Write(buffer2, 0, 8);
            Thread.Sleep(1000);
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";
        }

        public void blenderoder17()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x03;
            buffer[2] = 0xA4;
            buffer[3] = 0x07;
            buffer[4] = 0x06;
            buffer[5] = 0xDD;
            buffer[6] = 0x5D;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(1000);
            //关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x03;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6B;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer2 = new Byte[8];
            buffer2[0] = 0xCC;
            buffer2[1] = 0x01;
            buffer2[2] = 0xA4;
            buffer2[3] = 0x07;
            buffer2[4] = 0x06;
            buffer2[5] = 0xDD;
            buffer2[6] = 0x5B;
            buffer2[7] = 0x02;
            serialPort1.Write(buffer2, 0, 8);
            Thread.Sleep(1000);
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";
        }

        public void blenderoder18()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x03;
            buffer[2] = 0xA4;
            buffer[3] = 0x08;
            buffer[4] = 0x07;
            buffer[5] = 0xDD;
            buffer[6] = 0x5F;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(1000);
            //关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x03;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6B;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer2 = new Byte[8];
            buffer2[0] = 0xCC;
            buffer2[1] = 0x01;
            buffer2[2] = 0xA4;
            buffer2[3] = 0x08;
            buffer2[4] = 0x07;
            buffer2[5] = 0xDD;
            buffer2[6] = 0x5D;
            buffer2[7] = 0x02;
            serialPort1.Write(buffer2, 0, 8);
            Thread.Sleep(1000);
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";
        }

        public void blenderoder19()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x03;
            buffer[2] = 0xA4;
            buffer[3] = 0x09;
            buffer[4] = 0x08;
            buffer[5] = 0xDD;
            buffer[6] = 0x61;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(1000);
            //关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x03;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6B;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer2 = new Byte[8];
            buffer2[0] = 0xCC;
            buffer2[1] = 0x01;
            buffer2[2] = 0xA4;
            buffer2[3] = 0x09;
            buffer2[4] = 0x08;
            buffer2[5] = 0xDD;
            buffer2[6] = 0x5F;
            buffer2[7] = 0x02;
            serialPort1.Write(buffer2, 0, 8);
            Thread.Sleep(1000);

            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";
        }

        public void blenderoder20()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x03;
            buffer[2] = 0xA4;
            buffer[3] = 0x0A;
            buffer[4] = 0x09;
            buffer[5] = 0xDD;
            buffer[6] = 0x63;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(1000);
            //关断
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x03;
            buffer1[2] = 0xB4;
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x6B;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer2 = new Byte[8];
            buffer2[0] = 0xCC;
            buffer2[1] = 0x01;
            buffer2[2] = 0xA4;
            buffer2[3] = 0x0A;
            buffer2[4] = 0x09;
            buffer2[5] = 0xDD;
            buffer2[6] = 0x61;
            buffer2[7] = 0x02;
            serialPort1.Write(buffer2, 0, 8);
            Thread.Sleep(1000);

            Byte[] buffer3 = new Byte[8];
            buffer3[0] = 0xCC;
            buffer3[1] = 0x01;
            buffer3[2] = 0xB4;
            buffer3[3] = 0x0A;
            buffer3[4] = 0x01;
            buffer3[5] = 0xDD;
            buffer3[6] = 0x69;
            buffer3[7] = 0x02;
            serialPort1.Write(buffer3, 0, 8);
            Thread.Sleep(1000);

            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                info += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }

            //截取
            string str1 = info.Replace(" ", "");
            //str2是截取好的数据
            string str2 = str1.Substring(4, 2);//截取str1的1前两个字符

            switch (str2)
            {
                case "FE":
                    MessageBox.Show("正常执行"); break;
                case "FF":
                    MessageBox.Show("未知错误"); break;
                case "06":
                    MessageBox.Show("未知位置"); break;
                case "05":
                    MessageBox.Show("电机堵转"); break;
                case "04":
                    MessageBox.Show("电机忙"); break;
                case "03":
                    MessageBox.Show("光耦错误"); break;
                case "02":
                    MessageBox.Show("参数错误"); break;
                case "01":
                    MessageBox.Show("桢错误"); break;
                case "00":
                    MessageBox.Show("正常"); break;
                default:
                    MessageBox.Show("未返回任何命令"); break;

            }
            Thread.Sleep(2000);
            info = "";
        }

        //关断
       /* public void guanduan2()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x03;
            buffer[2] = 0xB4;
            buffer[3] = 0x0A;
            buffer[4] = 0x01;
            buffer[5] = 0xDD;
            buffer[6] = 0x6B;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
        }

        public void guanduan1()
        {
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x02;
            buffer[2] = 0xB4;
            buffer[3] = 0x0A;
            buffer[4] = 0x01;
            buffer[5] = 0xDD;
            buffer[6] = 0x6A;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
        }*/

        //复位
        /*public void resetxxx()
        {
            
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x00;
            buffer[2] = 0x45;
            buffer[3] = 0x00;
            buffer[4] = 0x00;
            buffer[5] = 0xDD;
            buffer[6] = 0xEE;
            buffer[7] = 0x01;
            serialPort1.Write(buffer, 0, 8);

            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x01;
            buffer1[2] = 0x45;
            buffer1[3] = 0x00;
            buffer1[4] = 0x00;
            buffer1[5] = 0xDD;
            buffer1[6] = 0xEF;
            buffer1[7] = 0x01;
            serialPort1.Write(buffer1, 0, 8);
            int length = serialPort1.BytesToRead;
            byte[] data = new byte[length];
            serialPort1.Read(data, 0, length);
            //对串口接收数据的处理，可对data进行解析
            Thread.Sleep(2000);
            for (int i = 0; i < length; i++)
            {
                string str = Convert.ToString(data[i], 16).ToUpper();
                infos += (str.Length == 1 ? "0" + str + " " : str + " ");//将接收到的数据以十六进制保存到变量
            }
            infos = "";
        }*/


        //发送报文内容方法
        public void starttest()
        {
            #region 发送检测报文
            Byte[] buffer = new Byte[8];
            buffer[0] = 0x16;
            buffer[1] = 0x06;
            buffer[2] = 0x00;
            buffer[3] = 0x00;
            buffer[4] = 0x30;
            buffer[5] = 0x00;
            buffer[6] = 0x9E;
            buffer[7] = 0xED;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(1000);
            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0x16;
            buffer1[1] = 0x06;
            buffer1[2] = 0x00;
            buffer1[3] = 0x01;
            buffer1[4] = 0x45;
            buffer1[5] = 0x44;
            buffer1[6] = 0xE9;
            buffer1[7] = 0x8E;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);
            Byte[] buffer2 = new Byte[8];
            buffer2[0] = 0x16;
            buffer2[1] = 0x03;
            buffer2[2] = 0x00;
            buffer2[3] = 0x0A;
            buffer2[4] = 0x00;
            buffer2[5] = 0x3D;
            buffer2[6] = 0xA7;
            buffer2[7] = 0x3E;
            serialPort1.Write(buffer2, 0, 8);
            #endregion
        }
       



        
        #region 按日期查询折线图数据
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            chart1.Series.Clear();
            series1.Points.Clear();
            series2.Points.Clear();
            series3.Points.Clear();
            series4.Points.Clear();
            series5.Points.Clear();
            series6.Points.Clear();
            chart1.ChartAreas[0].AxisX.Minimum = 1;

            conn.Open();
            // 获取日期选择器选中的日期
            DateTime selectedDate = dateTimePicker1.Value.Date;
            string aass = (selectedDate.ToString("yyyy-MM-dd"));
            
            comm = new MySqlCommand("select allyls,lanzao,lvzao,guizao,jiazao,yinzao from ain where DATE(dtimer)='" + aass + "'", conn);
            dr = comm.ExecuteReader(); /*查询*/

            // 添加折线
            // 添加折线
            //标记点边框颜色      
            series1.MarkerBorderColor = Color.Orange;
            //标记点边框大小
            series1.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series1.MarkerColor = Color.Orange;//AxisColor
            //标记点大小
            series1.MarkerSize = 8;
            //标记点类型     
            series1.MarkerStyle = MarkerStyle.Circle;
            series1.ChartType = SeriesChartType.Line;
            series1.Color = Color.Orange;
            series1.BorderWidth = 2;
            series1.IsValueShownAsLabel = false;
            series1.Name = "总叶绿素";
            //Series series2 = new Series();
            //标记点边框颜色      
            series2.MarkerBorderColor = Color.Blue;
            //标记点边框大小
            series2.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series2.MarkerColor = Color.Blue;//AxisColor
            //标记点大小
            series2.MarkerSize = 8;
            //标记点类型     
            series2.MarkerStyle = MarkerStyle.Circle;
            series2.ChartType = SeriesChartType.Line;
            series2.Color = Color.Blue;
            series2.BorderWidth = 1;
            series2.IsValueShownAsLabel = false;
            series2.Name = "蓝藻";
            //Series series3 = new Series();
            //标记点边框颜色      
            series3.MarkerBorderColor = Color.Green;
            //标记点边框大小
            series3.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series3.MarkerColor = Color.Green;//AxisColor
            //标记点大小
            series3.MarkerSize = 8;
            //标记点类型     
            series3.MarkerStyle = MarkerStyle.Circle;
            series3.ChartType = SeriesChartType.Line;
            series3.Color = Color.Green;
            series3.BorderWidth = 1;
            series3.IsValueShownAsLabel = false;
            series3.Name = "绿藻";
            //Series series4 = new Series();
            //标记点边框颜色      
            series4.MarkerBorderColor = Color.Gray;
            //标记点边框大小
            series4.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series4.MarkerColor = Color.Gray;//AxisColor
            //标记点大小
            series4.MarkerSize = 8;
            //标记点类型     
            series4.MarkerStyle = MarkerStyle.Circle;
            series4.ChartType = SeriesChartType.Line;
            series4.Color = Color.Gray;
            series4.BorderWidth = 1;
            series4.IsValueShownAsLabel = false;
            series4.Name = "硅藻";
            //Series series5 = new Series();
            //标记点边框颜色      
            series5.MarkerBorderColor = Color.Red;
            //标记点边框大小
            series5.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series5.MarkerColor = Color.Red;//AxisColor
            //标记点大小
            series5.MarkerSize = 8;
            //标记点类型     
            series5.MarkerStyle = MarkerStyle.Circle;
            series5.ChartType = SeriesChartType.Line;
            series5.Color = Color.Red;
            series5.BorderWidth = 1;
            series5.IsValueShownAsLabel = false;
            series5.Name = "甲藻";
            //Series series6 = new Series();
            //标记点边框颜色      
            series6.MarkerBorderColor = Color.Pink;
            //标记点边框大小
            series6.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series6.MarkerColor = Color.Pink;//AxisColor
            //标记点大小
            series6.MarkerSize = 8;
            //标记点类型     
            series6.MarkerStyle = MarkerStyle.Circle;
            series6.Color = Color.Pink;
            series6.BorderWidth = 1;
            series6.IsValueShownAsLabel = false;
            series6.ChartType = SeriesChartType.Line;
            series6.Name = "隐藻";

            chart1.Series.Add(series1);
            chart1.Series.Add(series2);
            chart1.Series.Add(series3);
            chart1.Series.Add(series4);
            chart1.Series.Add(series5);
            chart1.Series.Add(series6);

            // 添加数据点
            int i = 0;
            while (dr.Read())
            {
                series1.Points.AddXY(i, dr.GetDecimal("allyls"));
                series2.Points.AddXY(i, dr.GetDecimal("lanzao"));
                series3.Points.AddXY(i, dr.GetDecimal("lvzao"));
                series4.Points.AddXY(i, dr.GetDecimal("guizao"));
                series5.Points.AddXY(i, dr.GetDecimal("jiazao"));
                series6.Points.AddXY(i, dr.GetDecimal("yinzao"));
                i++;
            }

            dr.Close();
            conn.Close();
            
        }
        #endregion




        //刷新折线图
        #region 刷新折线图和下拉列表框及选择地址
        public void shuaxinzhexiantu()
        {
            chart1.Series.Clear();
            series1.Points.Clear();
            series2.Points.Clear();
            series3.Points.Clear();
            series4.Points.Clear();
            series5.Points.Clear();
            series6.Points.Clear();
            conn.Close();
            chartinfo();


        }

        //刷新下拉列表
        public void shuaxinxiala()
        {
            // 在这里重新加载下拉列表框的数据
            // 假设下拉列表框的名称为comboBox1
            //string connectionString = "your_mysql_connection_string_here";
            string query = "select DISTINCT  addres from ain";
            using (MySqlConnection connection = new MySqlConnection(strConn))
            {
                MySqlCommand command = new MySqlCommand(query, connection);
                connection.Open();
                using (MySqlDataReader reader = command.ExecuteReader())
                {
                    comboBox2.Items.Clear(); // 清空下拉列表框的数据
                    comboBox2.Text = "请选择地址";
                    while (reader.Read())
                    {
                        string itemText = reader.GetString(0); // 假设第一列是要显示的文本
                        comboBox2.Items.Add(itemText); // 将文本添加到下拉列表框中
                    }
                }
            }
            this.Refresh();
        }
        
        //下拉列表选择地址
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedAddress = comboBox2.SelectedItem.ToString();
            conn.Open();
            MySqlCommand comm = new MySqlCommand("select allyls,addres,dtimer,lanzao,lvzao,guizao,jiazao,yinzao,cdom,zhuodu,fo,fv,fm,fvfm,sigma,cn,wendu,dianya,allswl,lanswl,lvswl,guiswl,jiaswl,yinswl from ain WHERE addres LIKE '%" + selectedAddress + "%' limit 1", conn);
            dr = comm.ExecuteReader(); /*查询*/
            while (dr.Read())
            {
                label1.Text = dr.GetString("dtimer");
                label22.Text = dr.GetString("fvfm");
                label14.Text = dr.GetString("allswl");
                label48.Text = dr.GetString("allyls");
                textBox1.Text = dr.GetString("lanzao");
                textBox2.Text = dr.GetString("lvzao");
                textBox3.Text = dr.GetString("guizao");
                textBox4.Text = dr.GetString("jiazao");
                textBox5.Text = dr.GetString("yinzao");
                textBox15.Text = dr.GetString("fo");
                textBox14.Text = dr.GetString("fv");
                textBox13.Text = dr.GetString("fm");
                textBox12.Text = dr.GetString("sigma");
                textBox11.Text = dr.GetString("cn");
                textBox19.Text = dr.GetString("zhuodu");
                textBox18.Text = dr.GetString("cdom");
                textBox17.Text = dr.GetString("dianya");
                textBox16.Text = dr.GetString("wendu");
                textBox10.Text = dr.GetString("lanswl");
                textBox9.Text = dr.GetString("lvswl");
                textBox8.Text = dr.GetString("guiswl");
                textBox7.Text = dr.GetString("jiaswl");
                textBox6.Text = dr.GetString("yinswl");


            }
            dr.Close();
            conn.Close();
            conn.Close();
        }
        #endregion


        #region  单选框选择折线图显示条数
        private void materialRadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            chart1.Series.Clear();
            series1.Points.Clear();
            series2.Points.Clear();
            series3.Points.Clear();
            series4.Points.Clear();
            series5.Points.Clear();
            series6.Points.Clear();
            // 获取 Chart 控件的 X 轴
            Axis sxAxis = chart1.ChartAreas[0].AxisX;
            //chart1.ChartAreas[0].AxisX.Interval = 1;
            //chart1.ChartAreas[0].AxisX.Minimum = 1;
            //设置Y轴最小值
            //chart1.ChartAreas[0].AxisX.Minimum = 1;
            // 将 X 轴的 Minimum 属性设置为 0

            sxAxis.Minimum = 1;
            //修改折线图数据
            chart1.ChartAreas[0].AxisX.MajorGrid.LineDashStyle = ChartDashStyle.NotSet;
            chart1.ChartAreas[0].AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dash; //设置网格类型为虚线

            // 获取折线图的 Y 轴对象
            var yAxis = chart1.ChartAreas[0].AxisY;
            // 获取 Y 轴的刻度线对象，并设置其 LabelForeColor 属性为红色
            yAxis.MajorTickMark.LineColor = Color.Black;
            yAxis.LabelStyle.ForeColor = Color.Black;

            // 获取折线图的 X 轴对象
            var xAxis = chart1.ChartAreas[0].AxisX;
            // 获取 X 轴的刻度线对象，并设置其 LabelForeColor 属性为红色
            xAxis.MajorTickMark.LineColor = Color.Black;
            xAxis.LabelStyle.ForeColor = Color.Black;

            //折线图获取数据库值
            conn.Open();
            comm = new MySqlCommand("select allyls,lanzao,lvzao,guizao,jiazao,yinzao from ain limit 10", conn);
            dr = comm.ExecuteReader(); /*查询*/

            // 添加折线
            //series1.ToolTip = "当前数值：#VAL";
            //标记点边框颜色      
            series1.MarkerBorderColor = Color.Orange;
            //标记点边框大小
            series1.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series1.MarkerColor = Color.Orange;//AxisColor
            //标记点大小
            series1.MarkerSize = 8;
            //标记点类型     
            series1.MarkerStyle = MarkerStyle.Circle;
            series1.ChartType = SeriesChartType.Line;
            series1.Color = Color.Orange;
            series1.BorderWidth = 2;
            series1.IsValueShownAsLabel = false;
            series1.Name = "总叶绿素";
            //Series series2 = new Series();
            //标记点边框颜色      
            series2.MarkerBorderColor = Color.Blue;
            //标记点边框大小
            series2.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series2.MarkerColor = Color.Blue;//AxisColor
            //标记点大小
            series2.MarkerSize = 8;
            //标记点类型     
            series2.MarkerStyle = MarkerStyle.Circle;
            series2.ChartType = SeriesChartType.Line;
            series2.Color = Color.Blue;
            series2.BorderWidth = 2;
            series2.IsValueShownAsLabel = false;
            series2.Name = "蓝藻";
            //Series series3 = new Series();
            //标记点边框颜色      
            series3.MarkerBorderColor = Color.Green;
            //标记点边框大小
            series3.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series3.MarkerColor = Color.Green;//AxisColor
            //标记点大小
            series3.MarkerSize = 8;
            //标记点类型     
            series3.MarkerStyle = MarkerStyle.Circle;
            series3.ChartType = SeriesChartType.Line;
            series3.Color = Color.Green;
            series3.BorderWidth = 2;
            series3.IsValueShownAsLabel = false;
            series3.Name = "绿藻";
            //Series series4 = new Series();
            //标记点边框颜色      
            series4.MarkerBorderColor = Color.Gray;
            //标记点边框大小
            series4.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series4.MarkerColor = Color.Gray;//AxisColor
            //标记点大小
            series4.MarkerSize = 8;
            //标记点类型     
            series4.MarkerStyle = MarkerStyle.Circle;
            series4.ChartType = SeriesChartType.Line;
            series4.Color = Color.Gray;
            series4.BorderWidth = 2;
            series4.IsValueShownAsLabel = false;
            series4.Name = "硅藻";
            //Series series5 = new Series();
            //标记点边框颜色      
            series5.MarkerBorderColor = Color.Red;
            //标记点边框大小
            series5.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series5.MarkerColor = Color.Red;//AxisColor
            //标记点大小
            series5.MarkerSize = 8;
            //标记点类型     
            series5.MarkerStyle = MarkerStyle.Circle;
            series5.ChartType = SeriesChartType.Line;
            series5.Color = Color.Red;
            series5.BorderWidth = 2;
            series5.IsValueShownAsLabel = false;
            series5.Name = "甲藻";
            //Series series6 = new Series();
            //标记点边框颜色      
            series6.MarkerBorderColor = Color.Pink;
            //标记点边框大小
            series6.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series6.MarkerColor = Color.Pink;//AxisColor
            //标记点大小
            series6.MarkerSize = 8;
            //标记点类型     
            series6.MarkerStyle = MarkerStyle.Circle;
            series6.Color = Color.Pink;
            series6.BorderWidth = 2;
            series6.IsValueShownAsLabel = false;
            series6.ChartType = SeriesChartType.Line;
            series6.Name = "隐藻";

            chart1.Series.Add(series1);
            chart1.Series.Add(series2);
            chart1.Series.Add(series3);
            chart1.Series.Add(series4);
            chart1.Series.Add(series5);
            chart1.Series.Add(series6);

            // 添加数据点
            int i = 0;
            while (dr.Read())
            {
                series1.Points.AddXY(i, dr.GetDecimal("allyls"));
                series2.Points.AddXY(i, dr.GetDecimal("lanzao"));
                series3.Points.AddXY(i, dr.GetDecimal("lvzao"));
                series4.Points.AddXY(i, dr.GetDecimal("guizao"));
                series5.Points.AddXY(i, dr.GetDecimal("jiazao"));
                series6.Points.AddXY(i, dr.GetDecimal("yinzao"));
                i++;
            }

            dr.Close();
            conn.Close();
        }

        private void materialRadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            chart1.Series.Clear();
            series1.Points.Clear();
            series2.Points.Clear();
            series3.Points.Clear();
            series4.Points.Clear();
            series5.Points.Clear();
            series6.Points.Clear();

            chart1.ChartAreas[0].AxisX.Interval = 1;
            chart1.ChartAreas[0].AxisX.Minimum = 1;
            //修改折线图数据
            chart1.ChartAreas[0].AxisX.MajorGrid.LineDashStyle = ChartDashStyle.NotSet;
            chart1.ChartAreas[0].AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dash; //设置网格类型为虚线

            //折线图获取数据库值
            conn.Open();
            comm = new MySqlCommand("select allyls,lanzao,lvzao,guizao,jiazao,yinzao from ain limit 50", conn);
            dr = comm.ExecuteReader(); /*查询*/

            // 添加折线

            // 添加折线
            //标记点边框颜色      
            series1.MarkerBorderColor = Color.Orange;
            //标记点边框大小
            series1.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series1.MarkerColor = Color.Orange;//AxisColor
            //标记点大小
            series1.MarkerSize = 8;
            //标记点类型     
            series1.MarkerStyle = MarkerStyle.Circle;
            series1.ChartType = SeriesChartType.Line;
            series1.Color = Color.Orange;
            series1.BorderWidth = 2;
            series1.IsValueShownAsLabel = false;
            series1.Name = "总叶绿素";
            //Series series2 = new Series();
            //标记点边框颜色      
            series2.MarkerBorderColor = Color.Blue;
            //标记点边框大小
            series2.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series2.MarkerColor = Color.Blue;//AxisColor
            //标记点大小
            series2.MarkerSize = 8;
            //标记点类型     
            series2.MarkerStyle = MarkerStyle.Circle;
            series2.ChartType = SeriesChartType.Line;
            series2.Color = Color.Blue;
            series2.BorderWidth = 2;
            series2.IsValueShownAsLabel = false;
            series2.Name = "蓝藻";
            //Series series3 = new Series();
            //标记点边框颜色      
            series3.MarkerBorderColor = Color.Green;
            //标记点边框大小
            series3.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series3.MarkerColor = Color.Green;//AxisColor
            //标记点大小
            series3.MarkerSize = 8;
            //标记点类型     
            series3.MarkerStyle = MarkerStyle.Circle;
            series3.ChartType = SeriesChartType.Line;
            series3.Color = Color.Green;
            series3.BorderWidth = 2;
            series3.IsValueShownAsLabel = false;
            series3.Name = "绿藻";
            //Series series4 = new Series();
            //标记点边框颜色      
            series4.MarkerBorderColor = Color.Gray;
            //标记点边框大小
            series4.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series4.MarkerColor = Color.Gray;//AxisColor
            //标记点大小
            series4.MarkerSize = 8;
            //标记点类型     
            series4.MarkerStyle = MarkerStyle.Circle;
            series4.ChartType = SeriesChartType.Line;
            series4.Color = Color.Gray;
            series4.BorderWidth = 2;
            series4.IsValueShownAsLabel = false;
            series4.Name = "硅藻";
            //Series series5 = new Series();
            //标记点边框颜色      
            series5.MarkerBorderColor = Color.Red;
            //标记点边框大小
            series5.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series5.MarkerColor = Color.Red;//AxisColor
            //标记点大小
            series5.MarkerSize = 8;
            //标记点类型     
            series5.MarkerStyle = MarkerStyle.Circle;
            series5.ChartType = SeriesChartType.Line;
            series5.Color = Color.Red;
            series5.BorderWidth = 2;
            series5.IsValueShownAsLabel = false;
            series5.Name = "甲藻";
            //Series series6 = new Series();
            //标记点边框颜色      
            series6.MarkerBorderColor = Color.Pink;
            //标记点边框大小
            series6.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series6.MarkerColor = Color.Pink;//AxisColor
            //标记点大小
            series6.MarkerSize = 8;
            //标记点类型     
            series6.MarkerStyle = MarkerStyle.Circle;
            series6.Color = Color.Pink;
            series6.BorderWidth = 2;
            series6.IsValueShownAsLabel = false;
            series6.ChartType = SeriesChartType.Line;
            series6.Name = "隐藻";

            chart1.Series.Add(series1);
            chart1.Series.Add(series2);
            chart1.Series.Add(series3);
            chart1.Series.Add(series4);
            chart1.Series.Add(series5);
            chart1.Series.Add(series6);

            // 添加数据点
            int i = 0;
            while (dr.Read())
            {
                series1.Points.AddXY(i, dr.GetDecimal("allyls"));
                series2.Points.AddXY(i, dr.GetDecimal("lanzao"));
                series3.Points.AddXY(i, dr.GetDecimal("lvzao"));
                series4.Points.AddXY(i, dr.GetDecimal("guizao"));
                series5.Points.AddXY(i, dr.GetDecimal("jiazao"));
                series6.Points.AddXY(i, dr.GetDecimal("yinzao"));
                i++;
            }

            dr.Close();
            conn.Close();
        }

        private void materialRadioButton3_CheckedChanged(object sender, EventArgs e)
        {
            chart1.Series.Clear();
            series1.Points.Clear();
            series2.Points.Clear();
            series3.Points.Clear();
            series4.Points.Clear();
            series5.Points.Clear();
            series6.Points.Clear();

            chart1.ChartAreas[0].AxisX.Interval = 1;
            chart1.ChartAreas[0].AxisX.Minimum = 1;
            //修改折线图数据
            chart1.ChartAreas[0].AxisX.MajorGrid.LineDashStyle = ChartDashStyle.NotSet;
            chart1.ChartAreas[0].AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dash; //设置网格类型为虚线

            //折线图获取数据库值
            conn.Open();
            comm = new MySqlCommand("select allyls,lanzao,lvzao,guizao,jiazao,yinzao from ain limit 100", conn);
            dr = comm.ExecuteReader(); /*查询*/

            // 添加折线

            // 添加折线
            //标记点边框颜色      
            series1.MarkerBorderColor = Color.Orange;
            //标记点边框大小
            series1.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series1.MarkerColor = Color.Orange;//AxisColor
            //标记点大小
            series1.MarkerSize = 8;
            //标记点类型     
            series1.MarkerStyle = MarkerStyle.Circle;
            series1.ChartType = SeriesChartType.Line;
            series1.Color = Color.Orange;
            series1.BorderWidth = 2;
            series1.IsValueShownAsLabel = false;
            series1.Name = "总叶绿素";
            //Series series2 = new Series();
            //标记点边框颜色      
            series2.MarkerBorderColor = Color.Blue;
            //标记点边框大小
            series2.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series2.MarkerColor = Color.Blue;//AxisColor
            //标记点大小
            series2.MarkerSize = 8;
            //标记点类型     
            series2.MarkerStyle = MarkerStyle.Circle;
            series2.ChartType = SeriesChartType.Line;
            series2.Color = Color.Blue;
            series2.BorderWidth = 2;
            series2.IsValueShownAsLabel = false;
            series2.Name = "蓝藻";
            //Series series3 = new Series();
            //标记点边框颜色      
            series3.MarkerBorderColor = Color.Green;
            //标记点边框大小
            series3.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series3.MarkerColor = Color.Green;//AxisColor
            //标记点大小
            series3.MarkerSize = 8;
            //标记点类型     
            series3.MarkerStyle = MarkerStyle.Circle;
            series3.ChartType = SeriesChartType.Line;
            series3.Color = Color.Green;
            series3.BorderWidth = 2;
            series3.IsValueShownAsLabel = false;
            series3.Name = "绿藻";
            //Series series4 = new Series();
            //标记点边框颜色      
            series4.MarkerBorderColor = Color.Gray;
            //标记点边框大小
            series4.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series4.MarkerColor = Color.Gray;//AxisColor
            //标记点大小
            series4.MarkerSize = 8;
            //标记点类型     
            series4.MarkerStyle = MarkerStyle.Circle;
            series4.ChartType = SeriesChartType.Line;
            series4.Color = Color.Gray;
            series4.BorderWidth = 2;
            series4.IsValueShownAsLabel = false;
            series4.Name = "硅藻";
            //Series series5 = new Series();
            //标记点边框颜色      
            series5.MarkerBorderColor = Color.Red;
            //标记点边框大小
            series5.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series5.MarkerColor = Color.Red;//AxisColor
            //标记点大小
            series5.MarkerSize = 8;
            //标记点类型     
            series5.MarkerStyle = MarkerStyle.Circle;
            series5.ChartType = SeriesChartType.Line;
            series5.Color = Color.Red;
            series5.BorderWidth = 2;
            series5.IsValueShownAsLabel = false;
            series5.Name = "甲藻";
            //Series series6 = new Series();
            //标记点边框颜色      
            series6.MarkerBorderColor = Color.Pink;
            //标记点边框大小
            series6.MarkerBorderWidth = 3; //chart1.;// Xaxis 
            //标记点中心颜色
            series6.MarkerColor = Color.Pink;//AxisColor
            //标记点大小
            series6.MarkerSize = 8;
            //标记点类型     
            series6.MarkerStyle = MarkerStyle.Circle;
            series6.Color = Color.Pink;
            series6.BorderWidth = 2;
            series6.IsValueShownAsLabel = false;
            series6.ChartType = SeriesChartType.Line;
            series6.Name = "隐藻";

            chart1.Series.Add(series1);
            chart1.Series.Add(series2);
            chart1.Series.Add(series3);
            chart1.Series.Add(series4);
            chart1.Series.Add(series5);
            chart1.Series.Add(series6);

            // 添加数据点
            int i = 0;
            while (dr.Read())
            {
                series1.Points.AddXY(i, dr.GetDecimal("allyls"));
                series2.Points.AddXY(i, dr.GetDecimal("lanzao"));
                series3.Points.AddXY(i, dr.GetDecimal("lvzao"));
                series4.Points.AddXY(i, dr.GetDecimal("guizao"));
                series5.Points.AddXY(i, dr.GetDecimal("jiazao"));
                series6.Points.AddXY(i, dr.GetDecimal("yinzao"));
                i++;
            }

            dr.Close();
            conn.Close();
        }

        #endregion

        //重置按钮，当流程结束后才能点击清除文本框内容
        private void materialButton1_Click(object sender, EventArgs e)
        {
            textBox20.Text = string.Empty; textBox23.Text = string.Empty; textBox25.Text = string.Empty; textBox27.Text = string.Empty;
            textBox29.Text = string.Empty; textBox21.Text = string.Empty; textBox22.Text = string.Empty; textBox24.Text = string.Empty;
            textBox26.Text = string.Empty; textBox28.Text = string.Empty; textBox39.Text = string.Empty; textBox37.Text = string.Empty;
            textBox35.Text = string.Empty; textBox33.Text = string.Empty; textBox31.Text = string.Empty; textBox38.Text = string.Empty;
            textBox36.Text = string.Empty; textBox34.Text = string.Empty; textBox32.Text = string.Empty; textBox30.Text = string.Empty;

            pictureBox1.Image = null; pictureBox2.Image = null;pictureBox3.Image = null; pictureBox4.Image = null;pictureBox5.Image = null; pictureBox6.Image = null;
            pictureBox7.Image = null; pictureBox8.Image = null; pictureBox9.Image = null; pictureBox10.Image = null; pictureBox11.Image = null; pictureBox12.Image = null;
            pictureBox13.Image = null; pictureBox14.Image = null; pictureBox15.Image = null; pictureBox16.Image = null; pictureBox17.Image = null; pictureBox18.Image = null;
            pictureBox19.Image = null; pictureBox20.Image = null;
        }


        #region 限制样品文本框输入只允许输入汉字、英文、数字、退格
        private void textBox20_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox23_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox25_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox27_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox29_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox21_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox22_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox24_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox26_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox28_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox39_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox37_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox35_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox33_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox31_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox38_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox36_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox34_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox32_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }

        private void textBox30_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 只允许输入汉字、英文和数字
            // 只允许输入汉字、英文、数字、退格
            if (!(Char.IsLetterOrDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar)))
            {
                e.Handled = true; // 阻止输入
                return;
            }

            // 长度不超过20个字符
            if (textBox1.Text.Length >= 20 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true; // 阻止输入
                return;
            }
        }
        #endregion

        //终止按钮
        private void materialFloatingActionButton1_Click(object sender, EventArgs e)
        {
            shuaxinxiala();
            shuaxinzhexiantu();
            Thread.Sleep(2000);

            textBox20.Text = string.Empty; textBox23.Text = string.Empty; textBox25.Text = string.Empty; textBox27.Text = string.Empty;
            textBox29.Text = string.Empty; textBox21.Text = string.Empty; textBox22.Text = string.Empty; textBox24.Text = string.Empty;
            textBox26.Text = string.Empty; textBox28.Text = string.Empty; textBox39.Text = string.Empty; textBox37.Text = string.Empty;
            textBox35.Text = string.Empty; textBox33.Text = string.Empty; textBox31.Text = string.Empty; textBox38.Text = string.Empty;
            textBox36.Text = string.Empty; textBox34.Text = string.Empty; textBox32.Text = string.Empty; textBox30.Text = string.Empty;

            pictureBox1.Image = null; pictureBox2.Image = null; pictureBox3.Image = null; pictureBox4.Image = null; pictureBox5.Image = null; pictureBox6.Image = null;
            pictureBox7.Image = null; pictureBox8.Image = null; pictureBox9.Image = null; pictureBox10.Image = null; pictureBox11.Image = null; pictureBox12.Image = null;
            pictureBox13.Image = null; pictureBox14.Image = null; pictureBox15.Image = null; pictureBox16.Image = null; pictureBox17.Image = null; pictureBox18.Image = null;
            pictureBox19.Image = null; pictureBox20.Image = null;

            if (textBox20.Text!="" ||textBox23.Text != "" || textBox25.Text != "" || textBox27.Text != "" || textBox29.Text != ""
                    || textBox21.Text != "" || textBox22.Text != "" || textBox24.Text != "" || textBox26.Text != ""
                    || textBox28.Text != "" || textBox39.Text != "" || textBox37.Text != "" || textBox35.Text != ""
                    || textBox33.Text != "" || textBox31.Text != "" || textBox38.Text != "" || textBox36.Text != ""
                    || textBox34.Text != "" || textBox32.Text != "" || textBox30.Text != "")
            {

                MessageBox.Show("任务已终止");
            }
            else
            {
                MessageBox.Show("没有在执行的检测任务!");
            }
            
            textEndtrue();
            materialButton1.Enabled = true;
            materialButton4.Enabled = true;
        }

        private void chart1_GetToolTipText(object sender, ToolTipEventArgs e)
        {
            //判断鼠标是否移动到数据标记点，是则显示提示信息
            if (e.HitTestResult.ChartElementType == ChartElementType.DataPoint)
            {
                int i = e.HitTestResult.PointIndex;
                DataPoint dp = e.HitTestResult.Series.Points[i];
                //分别显示x轴和y轴的数值，其中{1:F3},表示显示的是float类型，精确到小数点后3位。                     
                string r = string.Format("数值:{0} ",  dp.YValues[0]);

                //鼠标相对于窗体左上角的坐标
                Point formPoint = this.PointToClient(Control.MousePosition);
                int x = formPoint.X;
                int y = formPoint.Y;
                //显示提示信息
                this.panel6.Visible = true;
                
                this.panel6.Location = new Point(x, y);
                this.label49.Text = r;
            }

            //鼠标离开数据标记点，则隐藏提示信息
            else
            {
                this.panel6.Visible = false;
            }
        }

       
    }
}
