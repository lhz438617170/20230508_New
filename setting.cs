using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CeBianLan
{
    internal class setting
    {
        
        private string cb7;//电源节能模式
        private string cb8;//开机自检模式
        private int tb40;//进样冲洗时间
        private int tb41;//进水冲洗时间
        private string cb12;//激光光源选择
        private int cb11; //蓝光初始光强
        private int cb14; //红光初始光强
        private int cb13;//PMT初始电压
        private int cb16;//PMT初始电压
        private string cb15;//初始增益等级
        private int tb42; //初始增益等级-0
        private float tb43;//初始增益等级-1
        private float tb44;//初始增益等级-2
        private float tb45;//初始增益等级-3
        private float tb46;//初始增益等级-4
        private float tb47;//空白样核查上限
        private int tb48;//标准样核查误差上限
        private long tb49;//蓝藻生物量系数
        private int tb50;//总叶绿素系数
        private float tb51;//蓝藻校正系数
        private long tb52;//绿藻生物量系数
        private int tb53;//绿藻校正系数
        private int tb54; //硅藻校正系数
        private long tb55;//硅藻生物量系数
        private int tb56;//甲藻校正系数
        private int tb57;//隐藻校正系数
        private long tb58;//甲藻生物量系数
        private int tb59;//CDOM校正系数
        private int tb60;//浊度校正系数
        private long tb61;//隐藻生物量系数

        public string Cb7 { get => cb7; set => cb7 = value; }
        public string Cb8 { get => cb8; set => cb8 = value; }
        public int Tb40 { get => tb40; set => tb40 = value; }
        public int Tb41 { get => tb41; set => tb41 = value; }
        public string Cb12 { get => cb12; set => cb12 = value; }
        public int Cb11 { get => cb11; set => cb11 = value; }
        public int Cb14 { get => cb14; set => cb14 = value; }
        public int Cb13 { get => cb13; set => cb13 = value; }
        public int Cb16 { get => cb16; set => cb16 = value; }
        public string Cb15 { get => cb15; set => cb15 = value; }
        public int Tb42 { get => tb42; set => tb42 = value; }
        public float Tb43 { get => tb43; set => tb43 = value; }
        public float Tb44 { get => tb44; set => tb44 = value; }
        public float Tb45 { get => tb45; set => tb45 = value; }
        public float Tb46 { get => tb46; set => tb46 = value; }
        public float Tb47 { get => tb47; set => tb47 = value; }
        public int Tb48 { get => tb48; set => tb48 = value; }
        public long Tb49 { get => tb49; set => tb49 = value; }
        public int Tb50 { get => tb50; set => tb50 = value; }
        public float Tb51 { get => tb51; set => tb51 = value; }
        public long Tb52 { get => tb52; set => tb52 = value; }
        public int Tb53 { get => tb53; set => tb53 = value; }
        public int Tb54 { get => tb54; set => tb54 = value; }
        public long Tb55 { get => tb55; set => tb55 = value; }
        public int Tb56 { get => tb56; set => tb56 = value; }
        public int Tb57 { get => tb57; set => tb57 = value; }
        public long Tb58 { get => tb58; set => tb58 = value; }
        public int Tb59 { get => tb59; set => tb59 = value; }
        public int Tb60 { get => tb60; set => tb60 = value; }
        public long Tb61 { get => tb61; set => tb61 = value; }
    }
}
