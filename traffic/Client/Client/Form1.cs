using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using VISSIM_COMSERVERLib;

namespace Client
{
    public partial class Form1 : Form
    {
        static Vissim vis;
        System.Threading.Thread MyThread;
        //声明一个数组，用来接受
        private List<int> recieviedList = new List<int>();
        private bool bInImg = false;
        private bool bInImg1 = false;
        private bool bInImg2 = false;
        private bool bInImg3 = false;
        private bool bInImg4 = false;
        private bool bInImg5 = false;

        TcpListener tcpListener;
        static List<String> recivedDataList = new List<string>();

        public Form1()
        {
            InitializeComponent();
        }
        //加载路网文件
        private void intialRoadNet() {
            vis = new Vissim();
            vis.LoadNet(@"E:\vissim4.3\Exe\node2.inp", 0);
            vis.Simulation.Resolution = 1;
            vis.Evaluation.set_AttValue("DATACOLLECTION", true); //软件激活datacollection检测器评价
            vis.Evaluation.set_AttValue("TRAVELTIME", true); //软件激活traveltime检测器评价
            vis.Evaluation.set_AttValue("QUEUECOUNTER", true);//软件激活queuecounter检测器评价
            vis.Evaluation.set_AttValue("DELAY", true);//软件激活dealy检测器评价


        }
        private void Form1_Load(object sender, EventArgs e)
        {
            Form1.CheckForIllegalCrossThreadCalls = false;
            MyThread = new System.Threading.Thread(new ThreadStart(ReceiveDataFromUDPClient));
            MyThread.IsBackground = true;
            MyThread.Start();
            pictureBox2.Image = imageList1.Images[1];
            pictureBox3.Image = imageList1.Images[3];
            pictureBox4.Image = imageList1.Images[5];
            pictureBox5.Image = imageList1.Images[1];
            pictureBox6.Image = imageList1.Images[3];
            pictureBox7.Image = imageList1.Images[5];

            timer1.Enabled = true;

        }
        public void ReceiveDataFromUDPClient()
        {
            while (true)
            {

                UdpClient udpClient = new UdpClient(1002);
                //传递到 ref 参数的参数RemoteIpEndPoint必须最先初始化
                IPEndPoint RemoteIpEndPoint = new IPEndPoint(IPAddress.Any, 0); // 括号里是应用程序连接到主机上的服务所需的主机和本地或远程端口信息
                Byte[] receiveBytes = udpClient.Receive(ref RemoteIpEndPoint);
                //接收到的receiveBytes转化为字符串
                string returnData = Encoding.Default.GetString(receiveBytes);
                //存到list中去
               
                //后台打印
                Console.WriteLine(returnData.ToString());
                //存入list集合
                recivedDataList.Add(returnData);
                //在textBox上显示
                textBox1.Text =returnData.ToString();
                //详情展示
                //找到目标字符的索引
                int a1 = returnData.IndexOf('D');
                int a2 = returnData.IndexOf('周');
                int a3 = returnData.IndexOf('绿');
                int a4 = returnData.IndexOf('红');
                int a5 = returnData.IndexOf('黄');
                
                //展示路口id
                textBox2.Text = returnData.Substring(a1+1,1);
                //展示周期
      
                textBox3.Text = returnData.Substring(a2+2,2);
                //直行相位
                //展示绿灯
                textBox4.Text = returnData.Substring(a3+2,a4-(a3+2));
                //展示红灯
                textBox5.Text = returnData.Substring(a4+2,a5-(a4+2));
                //展示黄灯
                textBox6.Text = returnData.Substring(a5+2);

                //右转相位
                //展示绿灯
                textBox7.Text = returnData.Substring(a4 + 2, a5 - (a4 + 2));
                //展示红灯
                textBox8.Text = returnData.Substring(a3 + 2, a4 - (a3 + 2));
                //展示黄灯
                textBox9.Text = returnData.Substring(a5 + 2);
                
                udpClient.Close();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        static void showRecivedData() {
            int count = recivedDataList.Count;
            foreach (String result in recivedDataList) {
                Console.WriteLine(result);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            showRecivedData();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {   //直行相位闪
            //绿灯闪
            if (bInImg == false)
            {
                pictureBox2.Image = imageList1.Images[1];
                bInImg = true;
            }
            else
            {
                pictureBox2.Image = imageList1.Images[0];
                bInImg = false;
            }
            //红灯闪
            if (bInImg1 == false)
            {
                pictureBox3.Image = imageList1.Images[3];
                bInImg1 = true;
            }
            else
            {
                pictureBox3.Image = imageList1.Images[2];
                bInImg1 = false;
            }
            //黄灯闪
            if (bInImg2 == false)
            {
                pictureBox4.Image = imageList1.Images[5];
                bInImg2 = true;
            }
            else
            {
                pictureBox4.Image = imageList1.Images[4];
                bInImg2 = false;
            }
            //右转相位闪
            //绿灯闪
            if (bInImg3 == false)
            {
                pictureBox7.Image = imageList1.Images[1];
                bInImg3 = true;
            }
            else
            {
                pictureBox7.Image = imageList1.Images[0];
                bInImg3 = false;
            }
            //红灯闪
            if (bInImg4 == false)
            {
                pictureBox6.Image = imageList1.Images[3];
                bInImg4 = true;
            }
            else
            {
                pictureBox6.Image = imageList1.Images[2];
                bInImg4 = false;
            }
            //黄灯闪
            if (bInImg5 == false)
            {
                pictureBox5.Image = imageList1.Images[5];
                bInImg5 = true;
            }
            else
            {
                pictureBox5.Image = imageList1.Images[4];
                bInImg5 = false;
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            intialRoadNet();

        }

    }
}
