using System;
using System.Drawing;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace 溫度計
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            notifyIcon1.Text = this.Text;
            notifyIcon1.Icon = this.Icon;
            this.ShowInTaskbar = false;

            this.zabbix_server_ip = System.Configuration.ConfigurationManager.AppSettings["zabbix_server_ip"];
            if (this.zabbix_server_ip != null)
            {
                this.zabbix_host = System.Configuration.ConfigurationManager.AppSettings["zabbix_host"];
                this.t1key = System.Configuration.ConfigurationManager.AppSettings["t1.key"];
                this.t2key = System.Configuration.ConfigurationManager.AppSettings["t2.key"];
                timer1.Interval = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["interval"]) * 1000;
                timer1.Enabled = true;
            }

            //获取当前工作区宽度和高度（工作区不包含状态栏）
            int ScreenWidth = Screen.PrimaryScreen.WorkingArea.Width;
            int ScreenHeight = Screen.PrimaryScreen.WorkingArea.Height;
            //计算窗体显示的坐标值，可以根据需要微调几个像素
            int x = ScreenWidth - this.Width - 5;
            int y = 5; // ScreenHeight - this.Height - 5;
            this.Location = new Point(x, y);
        }

        SerialPort serialPort = null;

        string zabbix_server_ip = "";
        string zabbix_host = "";
        string t1key = "";
        string t2key = "";

        string _t1 = "";
        string _t2 = "";

        public void InitSerialPort()
        {
            serialPort = new SerialPort();
            if (serialPort != null)
            {
                serialPort.PortName = System.Configuration.ConfigurationManager.AppSettings["port"];//端口号，这里可以电脑已经连接的COM口，如COM1;
                serialPort.DataReceived += SerialPort_DataReceived;
            }
            if (!serialPort.IsOpen)
            {
                serialPort.Open();//打开端口，进行监控
            }
        }

        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            string indata = sp.ReadExisting();
            Regex regex = new Regex(@"t([12]{1})=[+-]{1}([0-9.]+)", RegexOptions.IgnoreCase);
            Match m = regex.Match(indata);
            if (m.Success)
            {
                if (m.Groups[1].Value == "1")
                {
                    this._t1 = m.Groups[2].Value;
                }
                else
                {
                    this._t2 = m.Groups[2].Value;
                }

                this.Invoke((EventHandler)(delegate
                {
                    this.label4.Text = this._t1 + "°C";
                    this.label3.Text = this._t2 + "°C";
                }));
            }
        }

        [DllImport("shell32.dll ")]
        public static extern int ShellExecute(IntPtr hwnd, StringBuilder lpszOp, StringBuilder lpszFile, StringBuilder lpszParams, StringBuilder lpszDir, int FsShowCmd);

        private void call_zabbex_send(string t, string temp)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("-vv -z " + this.zabbix_server_ip);
            sb.Append(" -s " + this.zabbix_host);
            if (t == "t1")
                sb.Append(" -k " + this.t1key);
            else
                sb.Append(" -k " + this.t2key);
            sb.Append(" -o " + temp);
            ShellExecute(IntPtr.Zero, new StringBuilder("Open"), new StringBuilder(System.Configuration.ConfigurationManager.AppSettings["zabbix_sender"]), sb, new StringBuilder(""), 0);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            InitSerialPort();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.call_zabbex_send("t1", this._t1);
            this.call_zabbex_send("t2", this._t2);
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
        }

    }
}
