using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication1
{
    #region 帧格式
    public class EcgDataClass
    {
        public ushort Type;
        public ushort Length;
        public ushort[] LeadiValue = new ushort[96];//一维数组，最大容量为96
    };
    public class MotDataClass
    {
        public ushort Type;
        public ushort Length;
        public short[] xValue = new short[32];//三轴数据类，三类数据，三个电流320R、ACS712、MAX4372
        public short[] yValue = new short[32];//一维数组，最大容量为32
        public short[] zValue = new short[32];

    };
    public class OtherDataClass
    {
        public ushort Type;
        public ushort Length;
        public ushort Value;//其他非数组数据类
    };
    public class PackClass//数据包格式规则类
    {
        public ushort CodeVersion;
        public ushort Length;
        public ushort Crc;
        public ushort Id;
        public EcgDataClass EcgDataClass = new EcgDataClass();//心电信号
        public OtherDataClass EcgMarkClass = new OtherDataClass();//打标数据
        public MotDataClass AccDataClass = new MotDataClass();//加速度
        public MotDataClass GyrDataClass = new MotDataClass();//角速度
        public MotDataClass MagDataClass = new MotDataClass();//磁力计
        public OtherDataClass BatDataClass = new OtherDataClass();//电量
        public OtherDataClass ErrDataClass = new OtherDataClass();//错误信息
    }
    #endregion

    #region 一些宏定义
    class Global
    {
        public int AxisNum = 9;
        public int HeadSize = 8;
        public int AxisSize = 32;
        public int EcgSize = 96;
        public int PackSize = 810;
        public int CodeVersion_Index = 0;
        public int Length_Index = 2;
        public int Crc_Index = 4;
        public int Id_Index = 6;
        public int EcgData_Index = 8;
        public int EcgMark_Index = 204;
        public int AccData_Index = 210;
        public int GyrData_Index = 406;
        public int MagData_Index = 602;
        public int BatData_Index = 798;
        public int ErrData_Index = 804;
    }
    #endregion
}
