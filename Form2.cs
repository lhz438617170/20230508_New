using MySql.Data.MySqlClient;
using NPOI.SS.Formula.Functions;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace CeBianLan
{
    
    public partial class Form2 : Form
    {
        private DataTable dataTable;
        private int pageSize = 10;
        private int currentPage = 1;
        private int totalPage;
        string strConn = "";
        public Form2()
        {
            InitializeComponent();
            // 创建数据库连接器并打开连接
/*
            // 建立数据库连接
            SqlConnection conn = new SqlConnection("Database = hz_test; Server = localhost; Port = 3306; Password = root; UserID = root");
            conn.Open();

            // 定义 SQL 查询
            string query = "SELECT allyls FROM ain";

            // 执行查询
            SqlCommand cmd = new SqlCommand(query, conn);
            SqlDataReader reader = cmd.ExecuteReader();

            // 将查询结果添加到表格控件中
            while (reader.Read())
            {
                dataGridView1.Rows.Add(reader[0].ToString());
            }

            // 关闭数据库连接
            reader.Close();
            conn.Close();*/
            /*dataTable = new DataTable();
            dataTable.Columns.Add("ID", typeof(int));
            dataTable.Columns.Add("Name", typeof(string));

            // 添加数据到 DataTable
            for (int i = 1; i <= 10; i++)
            {
                dataTable.Rows.Add(i, $"Name {i}");
            }*/
            // 循环读取查询结果并将其添加到 DataTable 对象中

            /*textBox1.Text = " 16 06 00 00 30 00 9E ED 16 06 00 01 45 44 E9 8E 16 03 7A 09 47 41 BC E6 66 41 E3 A7 20 3E 13 00 00 00 00 00 00 00 00 A7 20 3E 13 00 00 00 00 00 00 00 00 48 77 3F C3 00 00 00 00 3D 83 44 34 00 00 00 00 00 00 00 00 3D 83 44 34 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 92 1A 45 85 53 82 45 C1 05 9F 44 EF 41 3E 3E 9E 0B 6F 35 52 B8 19 3F 4D 80 00 46 A2 00 00 00 00 00 00 00 00 00 00 00 00 00 00 03 43  ";

            inf = textBox1.Text;
            //去空格
            inf=inf.Replace(" ", "");
            //textBox1.Text = inf;
            //截取
            string str1 = inf;
            //str2是截取好的数据
            string str2 = str1.Substring(38, 248);//截取str1的1前两个字符

            string input = str2;
            string[] output = Enumerable.Range(0, input.Length / 8)
            .Select(i => input.Substring(i * 8, 8))
            .ToArray();


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
            string inp22= output[22];         //这是第一组的8个字符数据
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

            //formattedNum是最后得到的数据
            textBox1.Text = formattedNum.ToString()+"|"+formattedNum1.ToString() + "|" + formattedNum2.ToString()
                + "|" + formattedNum3.ToString() + "|" + formattedNum4.ToString() + "|" + formattedNum5.ToString()
                + "|" + formattedNum6.ToString() + "|" + formattedNum7.ToString() + "|" + formattedNum8.ToString()
                + "|" + formattedNum9.ToString() + "|" + formattedNum10.ToString() + "|" + formattedNum11.ToString()
                + "|" + formattedNum12.ToString() + "|" + formattedNum13.ToString() + "|" + formattedNum14.ToString()
                + "|" + formattedNum15.ToString() + "|" + formattedNum16.ToString() + "|" + formattedNum17.ToString()
                + "|" + formattedNum18.ToString() + "|" + formattedNum19.ToString() + "|" + formattedNum20.ToString()
                + "|" + formattedNum21.ToString() + "|" + formattedNum22.ToString() + "|" + formattedNum23.ToString()
                + "|" + formattedNum24.ToString() + "|" + formattedNum25.ToString() + "|" + formattedNum26.ToString()
                + "|" + formattedNum27.ToString() + "|" + formattedNum28.ToString() + "|" + formattedNum29.ToString();*/
        }

       

        private void button1_Click(object sender, EventArgs e)
        {
            serialPort1.Open();
            Byte[] buffer2 = new Byte[8];
            buffer2[0] = 0xCC;
            buffer2[1] = 0x02;
            buffer2[2] = 0xA4;
            buffer2[3] = 0x01;
            buffer2[4] = 0x0A;
            buffer2[5] = 0xDD;
            buffer2[6] = 0x5A;
            buffer2[7] = 0x02;
            serialPort1.Write(buffer2, 0, 8);
            Thread.Sleep(3000);
            //电机关断
            Byte[] buffer3 = new Byte[8];
            buffer3[0] = 0xCC;
            buffer3[1] = 0x02;
            buffer3[2] = 0xB4;
            buffer3[3] = 0x0A;
            buffer3[4] = 0x01;
            buffer3[5] = 0xDD;
            buffer3[6] = 0x6A;
            buffer3[7] = 0x02;
            serialPort1.Write(buffer3, 0, 8);
            Thread.Sleep(1000);

            //电机打开
            Byte[] buffer = new Byte[8];
            buffer[0] = 0xCC;
            buffer[1] = 0x01;
            buffer[2] = 0xA4;
            buffer[3] = 0x01;
            buffer[4] = 0x0A;
            buffer[5] = 0xDD;              //CC 01 A4 01 0A DD 59 02
            buffer[6] = 0x59;
            buffer[7] = 0x02;
            serialPort1.Write(buffer, 0, 8);
            Thread.Sleep(3000);

            Byte[] buffer1 = new Byte[8];
            buffer1[0] = 0xCC;
            buffer1[1] = 0x01;
            buffer1[2] = 0xB4;               //CC 01 B4 0A 01 DD 69 02
            buffer1[3] = 0x0A;
            buffer1[4] = 0x01;
            buffer1[5] = 0xDD;
            buffer1[6] = 0x69;
            buffer1[7] = 0x02;
            serialPort1.Write(buffer1, 0, 8);
            Thread.Sleep(1000);
            serialPort1.Close();


            /*string s1 = "cc02fe0000dda902cc02fe0000dda902";
            string s2 = s1.Substring(4, 2);
            //MessageBox.Show(s2);

            switch (s2) 
            {
                case "fe":
                    MessageBox.Show("正常");
                    break;
                case "02":
                    MessageBox.Show("错误");
                    break;
            }*/
            /*string str = "1";
            str += str;
            MessageBox.Show(str);*/

            /*StringBuilder sb = new StringBuilder(2);
            sb.Append("aa");
            sb.Append("cc");
            string ins= sb.ToString();
            MessageBox.Show(ins);

            string ii = "hello ";
            string ss= $"{ii}fuck you";

            MessageBox.Show(ss);*/


            //StringBuffer buffer = new StringBuffer("Hello World");


        }

        private void Form2_MouseDown(object sender, MouseEventArgs e)
        {

            
        }

        private void Form2_MouseUp(object sender, MouseEventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            
            /* // 计算总页数
             totalPage = (int)Math.Ceiling((double)dataTable.Rows.Count / pageSize);

             // 显示第一页数据
             BindData(1);*/

            DateTime startDate = dateTimePicker1.Value;
            DateTime endDate = dateTimePicker2.Value;
        }
        private void BindData(int page)
        {
            /*// 根据当前页数计算起始位置和结束位置
            int start = (page - 1) * pageSize;
            int end = Math.Min(start + pageSize, dataTable.Rows.Count);

            // 筛选 DataTable 中的数据
            DataTable filteredTable = dataTable.Clone();
            for (int i = start; i < end; i++)
            {
                filteredTable.ImportRow(dataTable.Rows[i]);
            }

            // 绑定数据到 DataGridView 控件
            dataGridView1.DataSource = filteredTable;

            // 更新分页控件
            label1.Text = $"Page {page} of {totalPage}";
            button2.Enabled = (page > 1);
            button3.Enabled = (page < totalPage);
            currentPage = page;*/
        }
        // 更新分页信息

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            BindData(currentPage - 1);
        }

        private void bindingNavigatorMoveNextItem_Click(object sender, EventArgs e)
        {
            // 更新 DataGridView 控件的显示位置
           
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            BindData(currentPage + 1);
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            /*Properties.Settings.Default.textBox5 = textBox5.Text;
            Properties.Settings.Default.Save();*/
            
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            
        }
    }


}
