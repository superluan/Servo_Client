using System;
using System.IO;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Diagnostics;
using System.Threading;//多线程
using System.IO.Ports;
using System.Timers;
using System.Windows.Forms.DataVisualization.Charting;
using Aspose.Cells;
//using MyNodeLink;
//using DrawGraph;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form//分布式类
    {
        #region 变量定义
        //public string chartTitles = "信号", SeriesName1 = "CH1", SeriesName2 = "CH2", SeriesName3 = "CH3";
        //PackClass PackClass = new PackClass();
       // Global Global = new Global();//实例化Global的对象
        //
        //private bool handerListening = false;
        ///private bool comClosing = false;
        //public bool[] tabControl_index = { false, false, false, false, false, false };//标签控制
        private Int32 SendLength, RcveLength, RcveOffset, RcveMode;
        private byte RcveTemp;
        private Queue<double> dataQueue = new Queue<double>(10);
        private Queue<double> dataQueue1 = new Queue<double>(10);
        private Queue<double> dataQueue2 = new Queue<double>(10);
        private Queue<double> dataQueue3 = new Queue<double>(10);
        //private Queue<string> dataQueue = new Queue<string>(4096);
        private int num = 10;//每次删除增加几个点
        public string str;
        int n = 0;//n就是表格的行数，用来决定每次表格++
        #endregion
        #region 窗体初始化
        public Form1()//窗体
        {
            InitializeComponent();//.NET平台自动初始化
            serialPort1.Encoding = Encoding.GetEncoding("GB2312");
            Control.CheckForIllegalCrossThreadCalls = false;
        }
        #endregion
        #region comboBox检测和CheckBox反选操作
        private void comboBox_SelectedIndexChanged(object sender, EventArgs e)//检测下拉框，不指定某个comboBox
        {
            ComboBox cbx = sender as ComboBox;
            if (cbx != null)
            {
                serialPort_config(1);
            }
        }
      

        private void checkBox_MouseUp(object sender, MouseEventArgs e)//检测，反转选择框状态checkBox
        {
            CheckBox cbx = sender as CheckBox;
            if (cbx.Checked == true)
            {
                cbx.Checked = false;
            }
            else
            {
                cbx.Checked = true;
            }
        }
        #endregion
        #region 保存串口设置，以xml形式储存
        private void serialPort_config(int rw)//形参是串口开关标志，设立存储数据文件的名字为serialPort.xml（可扩展标记语言）
        {
            string strFilename = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "serialPort.xml");
            if (rw == 0)//串口关
            {
                if (File.Exists(strFilename))//文件操作，用于保存本次的串口下拉框的选中值，下次开启默认上一次的值。666
                {
                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.Load(strFilename);
                    XmlNode xn = xmlDoc.SelectSingleNode("Node1");
                    XmlNodeList xnl = xn.ChildNodes;
                    foreach (XmlNode xnf in xnl)
                    {
                        XmlElement xe = (XmlElement)xnf;
                        switch (xe.Name)
                        {
                            case "PortName": comboBox1.SelectedIndex = Convert.ToInt32(xe.GetAttribute("Index")); break;
                            case "BaudRate": comboBox2.SelectedIndex = Convert.ToInt32(xe.GetAttribute("Index")); break;
                            case "Parity":   comboBox3.SelectedIndex = Convert.ToInt32(xe.GetAttribute("Index")); break;
                            case "DataBits": comboBox4.SelectedIndex = Convert.ToInt32(xe.GetAttribute("Index")); break;
                            case "StopBits": comboBox5.SelectedIndex = Convert.ToInt32(xe.GetAttribute("Index")); break;
                        }
                    }
                }
            }
            else
            {   //在文件中写入当前串口的所有信息
                XmlTextWriter xmlWriternew = new XmlTextWriter(strFilename, Encoding.Default);
                xmlWriternew.Formatting = Formatting.Indented;
                xmlWriternew.WriteStartDocument();
                xmlWriternew.WriteStartElement("Node1");        
                xmlWriternew.WriteStartElement("PortName");
                xmlWriternew.WriteAttributeString("Index", comboBox1.SelectedIndex.ToString());
                xmlWriternew.WriteAttributeString("Text", comboBox1.Text);
                xmlWriternew.WriteEndElement();
                xmlWriternew.WriteStartElement("BaudRate");
                xmlWriternew.WriteAttributeString("Index", comboBox2.SelectedIndex.ToString());
                xmlWriternew.WriteAttributeString("Text", comboBox2.Text);
                xmlWriternew.WriteEndElement();
                xmlWriternew.WriteStartElement("Parity");
                xmlWriternew.WriteAttributeString("Index", comboBox3.SelectedIndex.ToString());
                xmlWriternew.WriteAttributeString("Text", comboBox3.Text);
                xmlWriternew.WriteEndElement();
                xmlWriternew.WriteStartElement("DataBits");
                xmlWriternew.WriteAttributeString("Index", comboBox4.SelectedIndex.ToString());
                xmlWriternew.WriteAttributeString("Text", comboBox4.Text);
                xmlWriternew.WriteEndElement();
                xmlWriternew.WriteStartElement("StopBits");
                xmlWriternew.WriteAttributeString("Index", comboBox5.SelectedIndex.ToString());
                xmlWriternew.WriteAttributeString("Text", comboBox5.Text);
                xmlWriternew.WriteEndElement();
                xmlWriternew.WriteEndElement();
                xmlWriternew.Close();
            }
            serialPort1_state();//串口状态
        }
#endregion
        #region Form1_Load
        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;//初始索引值
            comboBox2.SelectedIndex = 6;
            comboBox3.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;
            comboBox5.SelectedIndex = 0;

            comboBox1.SelectedIndexChanged += new System.EventHandler(comboBox_SelectedIndexChanged);//若下拉框索引改变
            comboBox2.SelectedIndexChanged += new System.EventHandler(comboBox_SelectedIndexChanged);
            comboBox3.SelectedIndexChanged += new System.EventHandler(comboBox_SelectedIndexChanged);
            comboBox4.SelectedIndexChanged += new System.EventHandler(comboBox_SelectedIndexChanged);
            comboBox5.SelectedIndexChanged += new System.EventHandler(comboBox_SelectedIndexChanged);

            checkBox1.MouseUp += new System.Windows.Forms.MouseEventHandler(checkBox_MouseUp);//同上
            checkBox2.MouseUp += new System.Windows.Forms.MouseEventHandler(checkBox_MouseUp);
            checkBox3.MouseUp += new System.Windows.Forms.MouseEventHandler(checkBox_MouseUp);
           //初始化参数
            serialPort_config(0);
            SendLength = 0;
            RcveLength = 0;
            RcveOffset = 0;
            RcveMode = 0;
            InitChart();//进入即绘图
            this.timer1.Start();
            this.timer2.Start();
            //pictureBox1.Paint += pictureBox1_Paint;
            //NodeList = new MyNodeLink.Link<byte>(pictureBox1.Width);//像素宽度单位是像素。
            //封装 GDI+ 包含图形图像和其属性的像素数据的位图。 一个 Bitmap 是用来处理图像像素数据所定义的对象。
           // RTGBuffer = new Bitmap(pictureBox1.Width, pictureBox1.Height);//获取picturebox的宽度和高度。
           // RTG = Graphics.FromImage(RTGBuffer);//从指定的RTGbuffer创建新的Graphics,RTGbuffer是数据，见上句。
           // RTG.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;//RTG=Graphics。
           // RTGPen = new Pen(Color.Blue);//画笔颜色

            //AdjustableArrowCap aac;//定义箭头帽
            //acc = new System.Drawing.Drawing2D.AdjustableArrowCap();
            //Rp.CustomStartCap = aac;  //开始端箭头帽  
            //RTG.DrawLine(Rp, 20, 10, 20, 150);//坐标  
            //RTG.DrawLine(Rp, 280, 150, 20, 150);//横坐标 
            /*****************listview**************************/
            this.listView1.BeginUpdate();       //数据更新，UI暂时挂起，直到EndUpdate绘制控件，可以有效避免闪烁并大大提高加载速度
            listView1.View = View.Details;
            listView1.GridLines = true;
            listView1.Columns.Clear();
            listView1.Columns.Add("Time", 278, HorizontalAlignment.Center);
            listView1.Columns.Add("Position", 278, HorizontalAlignment.Center);
           // listView1.Columns.Add("Sample ID", 70, HorizontalAlignment.Center);
          //  listView1.Columns.Add("Measured Value", 150, HorizontalAlignment.Center);
            //listView1.Columns.Add("Temperature", 100, HorizontalAlignment.Center);

            ImageList iList = new ImageList();
            iList.ImageSize = new Size(1, 15);//宽度和高度值必须大于等于1且不超过256
            this.listView1.SmallImageList = iList;//这样的结果在第一列的前面多出了1个分量的宽，所有行的高度为24

            listView1.Refresh();
            this.listView1.EndUpdate();  //结束数据处理，UI界面一次性绘制。
        }
        #endregion

        public static int StrToInt(string str)
        {
            return int.Parse(str);
        }
        #region 串口扫描函数
        private void SearchAndAddSerialToComboBox(SerialPort MyPort, ComboBox MyBox)//成员函数，扫描可用的串口
        {                                                               //将可用端口号添加到ComboBox
            //string[] MyString = new string[20];                         //最多容纳20个，太多会影响调试效率
            string Buffer;                                              //缓存
            MyBox.Items.Clear();                                        //清空ComboBox内容
            for (int i = 1; i < 20; i++)                                //循环
            {
                try                                                     //核心原理是依靠try和catch完成遍历
                {
                    Buffer = "COM" + i.ToString();
                    MyPort.PortName = Buffer;
                    MyPort.Open();                                      //如果失败，后面的代码不会执行
                    MyBox.Items.Add(Buffer);                            //打开成功，添加至下俩列表
                    MyBox.Text = Buffer;
                    MyPort.Close();                                     //关闭

                }
                catch
                {

                }
            }
        }
        #endregion
        #region 委托事件
        //委托（delegate），有些类似函数指针，在需要使用回调函数时，都可以考虑使用Delegate
        delegate void SetricrichTextBox1Callback(string str);
        private void SetrichTextBox1(string str)
        {
            richTextBox1.AppendText(str);//把richTextBox1和现在加在一起后赋给 richTextBox1"相当于" + 号
            richTextBox1.Focus();//为控件设置输入焦点
        }

        delegate void SetChart1Callback(string str);
        private void SetChart1(string str)
        {
            int  position0;
            string[] strArray = str.Split(new char[] { ','});//用逗号分隔
            int[] intArray = new int[strArray.Length + 6];//BUG修复处，把strArray.Length改为5
            for (int i = 0; i < strArray.Length; i++)
            {
               int.TryParse(strArray[i], out intArray[i]);
            }

                dataQueue.Enqueue(intArray[1]/5); //AD8417入队
                dataQueue1.Enqueue(intArray[2]);    //电位器入队
                dataQueue2.Enqueue(intArray[3]);    //目标位置入队

            //数字显示代码
            double ian = ((double)intArray[1] / 4096) * 3.3 * 2.606;//max471采集电压
            double youxiao = ian * ((double)(intArray[5]) / 20000.0); 
            string repaly = intArray[2].ToString();
            string repaly1 = intArray[3].ToString();
            string pwm = intArray[5].ToString(); 
            int angle = (int)(((double)intArray[2] / 4096) * 360);//角度估计(10-350度)
            int temperature = intArray[4];//温度估计
            string ian0 = ian.ToString();
            string youxiao0 = youxiao.ToString();
            string angle0 = angle.ToString();
            string temperature0 = temperature.ToString();
            textBox6.Text = repaly;                                 //当前位置
            textBox1.Text = ian0;
            textBox9.Text = (intArray[1]/135).ToString();
            textBox8.Text = angle0;
            textBox5.Text = temperature0;
            textBox10.Text = pwm;
            textBox11.Text = repaly1;

            string n1 = Convert.ToString(n + 1, 10);
            position0 = intArray[2];
            string nian1 = Convert.ToString(position0, 10);

            ListViewItem lvi = new ListViewItem();
            lvi.ImageIndex = n;              //通过与imageList绑定，显示imageList中第i项图标
            this.listView1.BeginUpdate();   //数据更新，UI暂时挂起，直到EndUpdate绘制控件，可以有效避免闪烁并大大提高加载速度

            lvi.Text = n1;
            lvi.SubItems.Add(nian1);
            this.listView1.Items.Add(lvi);
            this.listView1.EndUpdate();  //结束数据处理，UI界面一次性绘制。

        }
        #endregion
        #region 状态栏显示
        private void serialPort1_state()//状态栏的COM状态显示
        {  
            toolStripStatusLabel3.Text = comboBox1.Text + "(" + comboBox2.Text+"," + comboBox3.Text+"," + comboBox4.Text+"," + comboBox5.Text + ")";
            if (serialPort1.IsOpen == true)
            {
                toolStripStatusLabel3.Text += "Opened";
            }
            else
            {
                toolStripStatusLabel3.Text += "Closed";    
            }
        }
        #endregion
        #region 串口接收数据函数
        private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)//串口接收数据
        {

            byte[] buffer = new byte[serialPort1.ReadBufferSize];//缓存区大小
            StringBuilder MyStringBuilder = new StringBuilder(serialPort1.ReadBufferSize * 2);//实际大小的二倍存储数据
           
            if (button1.Text == "打开串口")
            {
                return;//串口未打开
            }
            //三个参数含义：（被写入数据组，缓冲区数组中开始写入的偏移量，要读取的字节数）,返回值读取的字节数
            int length = serialPort1.Read(buffer, RcveOffset, buffer.Length - 1);
            if (checkBox2.Checked == true)//显示数据
            {
                if (checkBox1.Checked == true)//16进制显示模式，要进行转化再显示
                {
                    RcveOffset = 0;//偏移量为0
                    for (int i = 0; i < length; i++)
                    {
                        MyStringBuilder.Append(String.Format("{0:X2}", Convert.ToInt32(buffer[i])) + " ");
                    }
                    this.Invoke(new SetricrichTextBox1Callback(SetrichTextBox1), MyStringBuilder.ToString());
                    this.Invoke(new SetChart1Callback(SetChart1), MyStringBuilder.ToString());
                }
                else//即RcveMode == 0或checkBox1.Checked == false
                {
                    byte[] hz = new byte[2];
                    int j = 0;

                    if (RcveOffset == 1)
                    {
                        buffer[0] = RcveTemp;
                        length += 1;
                    }
                    for (int i = 0; i < length; i++)
                    {
                        if (buffer[i] < 0x80)
                        {
                            if (buffer[i] == '\n')
                            {
                                MyStringBuilder.Append("\r\n");
                            }
                            else
                            {
                                MyStringBuilder.Append(Convert.ToChar(buffer[i]));
                            }
                            j = 0;
                        }
                        else if (j == 0)
                        {
                            hz[0] = buffer[i];
                            j = 1;

                        }
                        else
                        {
                            hz[1] = buffer[i];
                            MyStringBuilder.Append(System.Text.Encoding.Default.GetString(hz));
                            j = 0;
                        }
                    }
                    if (j == 1)
                    {
                        RcveOffset = 1;
                        RcveTemp = buffer[length - 1];
                    }
                    else
                    {
                        RcveOffset = 0;
                    }
                    this.Invoke(new SetricrichTextBox1Callback(SetrichTextBox1), MyStringBuilder.ToString());
                    this.Invoke(new SetChart1Callback(SetChart1), MyStringBuilder.ToString());
                } 
            }
            RcveLength += length;//总接收数据计数
            toolStripStatusLabel2.Text = "接收" + " " + RcveLength.ToString("d");//状态栏显示总接收数
        }
        #endregion
        #region 串口打开后关闭选项
        private void ConsoleSerialPort(int n)//串口关闭和打开时对按钮和下拉栏的控制，防止错误操作
        {
            if (n == 0)
            {
                ovalShape1.FillColor = Color.Transparent;
                comboBox1.Enabled = true;
                comboBox2.Enabled = true;
                comboBox3.Enabled = true;
                comboBox4.Enabled = true;
                comboBox5.Enabled = true;
                button4.Enabled = false;
                button1.Text = "打开串口";
                Application.DoEvents();
                serialPort1.Close();
            }
            else
            {
                serialPort1.Open();
                ovalShape1.FillColor = Color.Chartreuse;
                comboBox1.Enabled = false;
                comboBox2.Enabled = false;
                comboBox3.Enabled = false;
                comboBox4.Enabled = false;
                comboBox5.Enabled = false;
                button4.Enabled = true;
                button1.Text = "关闭串口";
            }
            serialPort1_state();
        }
        #endregion
        #region 错误提示框
        [StructLayout(LayoutKind.Sequential)]
        public struct DEV_BROADCAST_HDR
        {
            public Int32 dbch_size;
            public Int32 dbch_devicetype;
            public Int32 dbch_reserved;
        }
        [StructLayout(LayoutKind.Sequential)]
        protected struct DEV_BROADCAST_PORT_Fixed
        {
            public uint dbcp_size;
            public uint dbcp_devicetype;
            public uint dbcp_reserved;
        }
        protected override void WndProc(ref Message m)//为了防止程序运行时拔出串口做的错误提示，获取系统提示信息加工弹框
        {
            if (m.Msg == 0x0219)
            {
                if (m.WParam.ToInt32() == 0x8004)
                {
                    DEV_BROADCAST_HDR dbhd = (DEV_BROADCAST_HDR)Marshal.PtrToStructure(m.LParam, typeof(DEV_BROADCAST_HDR));
                    string portName = Marshal.PtrToStringUni((IntPtr)(m.LParam.ToInt32() + Marshal.SizeOf(typeof(DEV_BROADCAST_PORT_Fixed))));
                    
                    if (dbhd.dbch_devicetype == 0x00000003)
                    {
                        if (portName == serialPort1.PortName)
                        {
                            if (button1.Text == "关闭串口")
                            {
                                ConsoleSerialPort(0);
                                MessageBox.Show("串口拨出", "串口调试助手", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                }
            }
            base.WndProc(ref m);
        }
#endregion
        #region 串口发送函数
        private void serialPort1_DataSend()//发送数据函数
        {
            if (checkBox3.Checked == true)//16进制格式发送
            {
                string SendText = textBox2.Text.Replace(" ", "");//替换逗号
                byte[] SendBuffer = new byte[SendText.Length / 2];//发送缓存区大小
                int Length = 0;
                for (int i = 0; i < SendText.Length / 2; i++)
                {
                    if (Regex.IsMatch(SendText.Substring(i * 2, 2), @"[0-9a-fA-F]{2,}$"))
                    {
                        SendBuffer[Length++] = Convert.ToByte(SendText.Substring(i * 2, 2), 16);
                    }
                    else
                    {
                        MessageBox.Show("字符格式:" + SendText.Substring(i * 2, 2), "串口调试助手", MessageBoxButtons.OK, MessageBoxIcon.Question);
                    }
                }
                serialPort1.Write(SendBuffer, 0, Length);//写数据
                SendLength += Length;//发送数据的长度
            }
            else//字符串发送
            {
                   byte[] SendBuffer = System.Text.Encoding.Default.GetBytes(textBox2.Text);
                    serialPort1.Write(SendBuffer, 0, SendBuffer.Length);
                    SendLength += SendBuffer.Length;
            }
            toolStripStatusLabel1.Text = "发送" + " " + SendLength.ToString("d");//状态栏
        }
        # endregion
        #region 窗口切换
        private void tabControl1_Selected(object sender, TabControlEventArgs e)//该控件生成两个可切换窗口
        {
            RcveMode = tabControl1.SelectedIndex;//接收模式为窗口选项的值
            if (RcveMode == 1)//绘图模式时
            {
                checkBox1.Enabled = false;//16进制显示按钮无效
                label6.Visible = true;
                label7.Visible = true;
                label9.Visible = true;
                label10.Visible = true;
                label26.Visible = true;
                label27.Visible = true;
            }
            else
            {
                checkBox1.Enabled = true;
                label6.Visible = false; 
                label7.Visible = false;             
                label9.Visible = false;
                label10.Visible = false;
                label26.Visible = false;
                label27.Visible = false;
            }
        }
        #endregion
        #region 串口接收Button事件集合
        private void button1_MouseUp(object sender, MouseEventArgs e)//打开串口，当鼠标指针在控件上并释放鼠标按键时发生。
        {
            if (serialPort1.IsOpen == true)
            {
                ConsoleSerialPort(0);
            }
            else
            {
                serialPort1.PortName = comboBox1.Text;
                serialPort1.BaudRate = Convert.ToInt32(comboBox2.Text);
                switch (comboBox3.Text)
                {
                    case "None":
                        serialPort1.Parity = System.IO.Ports.Parity.None;
                        break;
                    case "Even":
                        serialPort1.Parity = System.IO.Ports.Parity.Even;
                        break;
                    case "Odd":
                        serialPort1.Parity = System.IO.Ports.Parity.Odd;
                        break;
                }
                serialPort1.DataBits = Convert.ToInt32(comboBox4.Text);
                switch (comboBox5.Text)
                {
                    case "One":
                        serialPort1.StopBits = System.IO.Ports.StopBits.One;
                        break;
                    case "Two":
                        serialPort1.StopBits = System.IO.Ports.StopBits.Two;
                        break;
                    case "OnePointFive":
                        serialPort1.StopBits = System.IO.Ports.StopBits.OnePointFive;
                        break;
                }
                try
                {
                    ConsoleSerialPort(1);
                }
                catch (Exception ex) 
                {
                    MessageBox.Show(ex.Message, "串口调试助手", MessageBoxButtons.OK, MessageBoxIcon.Warning);      
                }
            }
        }

        private void button2_MouseUp(object sender, MouseEventArgs e)//清除窗口
        {
            RcveOffset = 0;
            richTextBox1.Text = "";
            this.listView1.Items.Clear();//清除listview
            n = 0;
            //NodeList = new MyNodeLink.Link<byte>(pictureBox1.Width);
        }
        private void button3_Click(object sender, EventArgs e)//保存数据键
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "txt files(*.txt)|*.txt";   // All files(*.*)|*.*
            saveFileDialog1.RestoreDirectory = true;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string localFileName = saveFileDialog1.FileName;
                StreamWriter sw = new StreamWriter(localFileName);
                sw.Write(richTextBox1.Text);
                sw.Close();
            }
        }
        private void button4_MouseUp(object sender, MouseEventArgs e)//发送数据
        {
            serialPort1_DataSend();
        }

        private void button5_MouseUp(object sender, MouseEventArgs e)//清空发送
        {
            textBox2.Text = "";
        }

        private void button6_MouseUp(object sender, MouseEventArgs e)//清空计数
        {
            SendLength = 0;
            RcveLength = 0;
            toolStripStatusLabel1.Text = "发送" + " " + SendLength.ToString("d");
            toolStripStatusLabel2.Text = "接收" + " " + RcveLength.ToString("d");
        }
#endregion
        #region 可视化窗口设置区Chart

        private void timer1_Tick(object sender, EventArgs e)//定时器中断，更新Chart
        {
            UpdateQueueValue();//更新dataQueue

            this.chart1.Series[0].Points.Clear();
            this.chart1.Series[1].Points.Clear();
            this.chart1.Series[2].Points.Clear();

            for (int i = 0; i < dataQueue.Count; i++)
            {
                this.chart1.Series[0].Points.AddXY((i+1), dataQueue.ElementAt(i));//依次将dataQueue中的点的XY坐标绘制
                this.chart1.Series[1].Points.AddXY((i+1), dataQueue1.ElementAt(i));
                this.chart1.Series[2].Points.AddXY((i+1), dataQueue2.ElementAt(i));
            }
        }
        private void button7_Click(object sender, EventArgs e)
        {
            this.timer1.Start();
            this.timer2.Start();
            InitChart();
        }
        private void button8_Click(object sender, EventArgs e)
        {
            this.timer1.Stop();
            this.timer2.Stop();

        }
        private void button9_Click(object sender, EventArgs e)
        {
            InitChart();
            this.timer1.Stop();
            this.timer2.Stop();
        }

        private void InitChart()
        {
            //定义图表区域
            this.chart1.ChartAreas.Clear();
            ChartArea chartArea1 = new ChartArea("C1");
            this.chart1.ChartAreas.Add(chartArea1);
            //定义存储和显示点的容器
            this.chart1.Series.Clear();
            Series series1 = new Series("S1");
            Series series2 = new Series("S2");
            Series series3 = new Series("S3");
            series1.ChartArea = "C1";
            series2.ChartArea = "C1";
            series3.ChartArea = "C1";
            this.chart1.Series.Add(series1);
            this.chart1.Series.Add(series2);
            this.chart1.Series.Add(series3);
            //设置图表显示样式，坐标轴和网格线
            this.chart1.ChartAreas[0].AxisY.Minimum = 0;
            this.chart1.ChartAreas[0].AxisY.Maximum = 4500;
            this.chart1.ChartAreas[0].AxisX.Interval = 5;
            this.chart1.ChartAreas[0].AxisX.MajorGrid.LineColor = System.Drawing.Color.Silver;
            this.chart1.ChartAreas[0].AxisY.MajorGrid.LineColor = System.Drawing.Color.Silver;
            //设置标题
            this.chart1.Titles.Clear();
            this.chart1.Titles.Add("S01");
            this.chart1.Titles[0].Text = "显示";
            this.chart1.Titles[0].ForeColor = Color.RoyalBlue;
            this.chart1.Titles[0].Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            //设置图表显示样式
            this.chart1.Series[0].Color = Color.Red;
            this.chart1.Series[1].Color = Color.Blue;
            this.chart1.Series[2].Color = Color.Green;

            if (rb1.Checked)
            {
                this.chart1.Titles[0].Text = string.Format("{0}显示",rb1.Text);
                this.chart1.Series[0].ChartType = SeriesChartType.Line;//线
                this.chart1.Series[1].ChartType = SeriesChartType.Line;//线
                this.chart1.Series[2].ChartType = SeriesChartType.Line;//线
            }
            if (rb2.Checked)
            {
                this.chart1.Titles[0].Text = string.Format("{0}显示",rb2.Text);
                this.chart1.Series[0].ChartType = SeriesChartType.Column;//柱状图
                this.chart1.Series[1].ChartType = SeriesChartType.Column;//柱状图
                this.chart1.Series[2].ChartType = SeriesChartType.Column;//柱状图
            }
            this.chart1.Series[0].Points.Clear();
            this.chart1.Series[1].Points.Clear();
            this.chart1.Series[2].Points.Clear();
        }

        private void UpdateQueueValue()
        {       
            if (dataQueue.Count >150)
            {
                //先出列，防止数据溢出
                for (int i = 0; i < num; i++)
                {
                    dataQueue.Dequeue();
                    dataQueue1.Dequeue();
                    dataQueue2.Dequeue();
                }
            }
            if (rb1.Checked)//进栈处理-折线
            {        
                this.chart1.Titles[0].Text = string.Format("{0}显示", rb1.Text);
                this.chart1.Series[0].ChartType = SeriesChartType.Line;//线
                this.chart1.Series[1].ChartType = SeriesChartType.Line;//线
                this.chart1.Series[2].ChartType = SeriesChartType.Line;//线

                //此段代码把String类型，转化为INT类数组，可以直接入队
                /* string x = "1,2,3,4,5,6,7";
                string[] strArray =x.Split(new char[] { ',' });               
                int[] intArray;
                intArray = Array.ConvertAll<string, int>(strArray, s => int.Parse(s));*/
              /*  for (int i = 0; i <6; i++)
                {
                   dataQueue.Enqueue(intArray[i]);
                }*/
              /*   Random r = new Random();
                for (int i = 0; i < 6; i++)
                {
                    dataQueue1.Enqueue(r.Next(0, 10));
                }*/
            }
            if (rb2.Checked)//进栈处理-柱状图
            {
                this.chart1.Titles[0].Text = string.Format("{0}显示", rb2.Text);
                this.chart1.Series[0].ChartType = SeriesChartType.Column;//柱状图
                this.chart1.Series[1].ChartType = SeriesChartType.Column;//柱状图
                this.chart1.Series[2].ChartType = SeriesChartType.Column;//柱状图
              /*  Random r = new Random();
                for (int i = 0; i < num; i++)
                {
                    dataQueue.Enqueue(r.Next(0, 10));
                }*/
               /* for (int i = 0; i < num; i++)
                {
                    //对curValue只取[0,360]之间的值
                    curValue = curValue % 360;
                    //对得到的正弦值，放大50倍，并上移50
                    dataQueue.Enqueue((500 * Math.Sin(curValue * Math.PI / 180)) + 1000);
                    curValue = curValue + 10;
                }*/
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            n++;
        }
        private void button11_Click(object sender, EventArgs e)//调用保存为Excel函数
        {
            List<int> list = new List<int>() { 5, 20, 12, 15, 15 };
            ReportToExcel(listView1, list, "测量数据表格");
        }
        #region Excel生成函数
        public void ReportToExcel(ListView list, List<int> ColumnWidth, string ReportTitleName)
        {
            //获取用户选择的excel文件名称
            string path;
            SaveFileDialog savefile = new SaveFileDialog();
            savefile.Filter = "Excel files(*.xls)|*.xls";
            if (savefile.ShowDialog() == DialogResult.OK)
            {
                //获取保存路径
                path = savefile.FileName;
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                Cells cell = ws.Cells;
                //定义并获取导出的数据源
                string[,] _ReportDt = new string[list.Items.Count, list.Columns.Count];
                for (int i = 0; i < list.Items.Count; i++)
                {
                    for (int j = 0; j < list.Columns.Count; j++)
                    {
                        _ReportDt[i, j] = list.Items[i].SubItems[j].Text.ToString();
                    }
                }
                //合并第一行单元格
                Range range = cell.CreateRange(0, 0, 1, list.Columns.Count);
                range.Merge();
                cell["A1"].PutValue(ReportTitleName); //标题

                //设置行高
                cell.SetRowHeight(0, 20);

                //设置字体样式
                Style style1 = wb.Styles[wb.Styles.Add()];
                style1.HorizontalAlignment = TextAlignmentType.Center;//文字居中
                style1.Font.Name = "宋体";
                style1.Font.IsBold = true;//设置粗体
                style1.Font.Size = 12;//设置字体大小

                Style style2 = wb.Styles[wb.Styles.Add()];
                style2.HorizontalAlignment = TextAlignmentType.Center;
                style2.Font.Size = 10;

                //给单元格关联样式
                cell["A1"].SetStyle(style1); //报表名字 样式

                //设置Execl列名
                for (int i = 0; i < list.Columns.Count; i++)
                {
                    cell[1, i].PutValue(list.Columns[i].Text);
                    cell[1, i].SetStyle(style2);
                }

                //设置单元格内容
                int posStart = 2;
                for (int i = 0; i < list.Items.Count; i++)
                {
                    for (int j = 0; j < list.Columns.Count; j++)
                    {
                        cell[i + posStart, j].PutValue(_ReportDt[i, j].ToString());
                        cell[i + posStart, j].SetStyle(style2);
                    }
                }

                //设置列宽
                for (int i = 0; i < list.Columns.Count; i++)
                {
                    cell.SetColumnWidth(i, Convert.ToDouble(ColumnWidth[i].ToString()));
                }
                //保存excel表格
                wb.Save(path);
            }
        }
        #endregion

        private void button10_Click(object sender, EventArgs e)
        {
            SearchAndAddSerialToComboBox(serialPort1, comboBox1);
            if (comboBox1.Text == "")
            {
                MessageBox.Show("未扫描到串口，请检查.", "错误");
            }
        }

        private void button11_Click_1(object sender, EventArgs e)
        {
            List<int> list = new List<int>() { 5, 20, 12, 15, 15 };
            ReportToExcel(listView1, list, "测量数据表格");
        }
    }
#endregion

}