using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using VISSIM_COMSERVERLib;
using System.Net.Sockets;
using System.Net;
using Newtonsoft.Json;

namespace vissim_ercikaifa
{
    public partial class Form1 : Form
    {
        
        static Vissim vis;
        public SignalGroup group;
        public SignalGroups groups;

        int itemTimeStep;
        bool startFlag; //仿真开始标志位
        public int qd1;//直行相位流量
        public int qd2;//右转相位流量
        public int sd;//道路饱和流量

        public int c0;//周期
        public int G;//直行相位绿灯
        public int R;//右转相位绿灯

        // public int quedingdianjicishu;//定义确定按钮点击次数
        public int[] infoArr = new int[20];//周期信息数组
        public int[] firstgreenArr = new int[20];//
        public int[] secondgreenArr = new int[20];

        //定义检测器检测时长，检测周期
        int detectPeriod = 20;
        int detectTime = 800;

        public int[,] dataCollectionVehiclesNum = new int[1000, 1000]; //定义检测器车流量记录数组
        public int[,] dataCollectionOccupancyRate = new int[1000, 1000]; //定义检测器密度记录数组
        public int[,] dataCollectionSpeed = new int[1000, 1000]; //定义检测器速度记录数组
                                                                 //行：检测器的编号，列：检测序列的数据
        public int[,] queueCounterLength = new int[1000, 1000];//定义排队检测器的记录数组（平均排队长度）
        public int[,] queueCounterStops = new int[1000, 1000];//定义排队检测器的记录数组（排队区域的停车次数）

        public int[,] delayDelay = new int[1000, 1000];
        public DataTable dtDataCollection = new DataTable(); //实例化DataCollection数据表对象
        public DataTable dcqueue = new DataTable();
        public DataTable dclink = new DataTable();
        public DataTable delayData = new DataTable();
        //DataCollections表DataGidView控件设置
        public bool fangzhen;
        public bool peishifinished;
        public bool fenshiduanFinished;
        public bool dingshiFinished;
        public int systemType;//用于区分系统的控制策略的参数，
        public bool pingjiaclicked;

        static String waitingforSend;

        private void columnAdd() {
            if (peishifinished || fenshiduanFinished||dingshiFinished) {
                //datacollection检测器的表头
                DataColumn dcDataCollectionTime = new DataColumn("检测时间");
                DataColumn dcDataCollectionVehicleNum = new DataColumn("车辆总数");
                DataColumn dcDataCollectionOcc = new DataColumn("占有率");
                DataColumn dcDataCollectionSpeed = new DataColumn("速度");
                ////queuecollection的表头
                DataColumn dcqueueTime = new DataColumn("检测时间");
                DataColumn dcqueueLength = new DataColumn("平均排队长度");
                DataColumn dcqueueStops = new DataColumn("停车次数");

                /////delayData的表头
                DataColumn dcDelayTime = new DataColumn("检测时间");
                DataColumn dcDelaydelay = new DataColumn("单车平均延误");
                //  DataColumn dcDelayStopped = new DataColumn("单车平均停车时间");


                dtDataCollection.Columns.Add(dcDataCollectionTime);
                dtDataCollection.Columns.Add(dcDataCollectionVehicleNum);
                dtDataCollection.Columns.Add(dcDataCollectionOcc);
                dtDataCollection.Columns.Add(dcDataCollectionSpeed);

                dcqueue.Columns.Add(dcqueueTime);
                dcqueue.Columns.Add(dcqueueLength);
                dcqueue.Columns.Add(dcqueueStops);


                delayData.Columns.Add(dcDelayTime);
                delayData.Columns.Add(dcDelaydelay);
                // delayData.Columns.Add(dcDelayStopped);
            }
        }
        //清空评价信息的表头和上一次存入的信息
        private void columnClear() {
            if (peishifinished || fenshiduanFinished) {
                dtDataCollection.Columns.Clear();
                dcqueue.Columns.Clear();
                delayData.Columns.Clear();
                dataTableClear(dataCollectionVehiclesNum);
                dataTableClear(dataCollectionOccupancyRate);
                dataTableClear(dataCollectionSpeed);
                dataTableClear(queueCounterLength);
                dataTableClear(queueCounterStops);
            }
        }
        private void dataTableClear(int[,] tableindex) {
            for (int i = 0; i < 200; i++) {
                for (int j = 0; j < 200; j++) {
                    tableindex[i, j] = 0;
                }
            }


        }
        public Form1()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
            // timer1.Start();
        }



        private void label3_Click(object sender, EventArgs e)
        {

        }
        //尝试着对周期进行改变，vissim中在一个运行周期中没有办法对cycletime的参数进行改变。
        public void TimerUpdateCycle() {
            //for(int i=0;i<4;i++)
            //while (itemTimeStep%infoArr[i]==0&&itemTimeStep!=0) {
            //        infoArr[i] = 30;
            //}
        }
        //加载路网
        private void InitialSystem() {
            vis = new Vissim();
            vis.LoadNet(@"E:\vissim4.3\Example\lianglukou.inp", 0);
            vis.Simulation.Resolution = 1;
            vis.Evaluation.set_AttValue("DATACOLLECTION", true); //软件激活datacollection检测器评价
            vis.Evaluation.set_AttValue("TRAVELTIME", true); //软件激活traveltime检测器评价
            vis.Evaluation.set_AttValue("QUEUECOUNTER", true);//软件激活queuecounter检测器评价
            vis.Evaluation.set_AttValue("DELAY", true);//软件激活dealy检测器评价

            //设置各检测器的评价周期


            label15.Text = ""; //仿真步数显示控件初始化
            progressBar1.Minimum = 0; //进度条显示控件最小值设置
            progressBar1.Maximum = Convert.ToInt32(vis.Simulation.Period); //进度条显示控件最大值设置
            //仿真信息表格展示
            VehicleInputs vehins = vis.Net.VehicleInputs;

            listView2.Clear();
            listView2.View = View.Details;
            listView2.FullRowSelect = true;
            listView2.Columns.Add("交叉口ID", 100, HorizontalAlignment.Center);
            listView2.Columns.Add("延时时间(s)", 100, HorizontalAlignment.Center);
            listView2.Columns.Add("仿真时间(s)", 100, HorizontalAlignment.Center);
            listView2.Columns.Add("输入交通量(pcu)", 100, HorizontalAlignment.Center);
            for (int i = 1; i <= vehins.Count; i++)
            {
                ListViewItem itm = listView2.Items.Add(vehins[i].ID.ToString());
                itm.SubItems.AddRange(new string[] { (vehins[i].AttValue["TIMEFROM"]).ToString(), vis.Simulation.Period.ToString(), (string)(vehins[i].AttValue["VOLUME"]).ToString() });
            }

            //修改vissim中各交叉口信号机周期

            for (int i = 1; i <= 4; i++)
            {
                vis.Net.SignalControllers.GetSignalControllerByNumber(i).AttValue["CYCLETIME"] = infoArr[i - 1];
                groups = vis.Net.SignalControllers[i].SignalGroups;

                for (int j = 1; j <= Convert.ToInt32(groups.Count); j++)
                {

                    group = groups.GetSignalGroupByNumber(j);
                    if (j == 1)
                    {
                        group.AttValue["GREENEND"] = firstgreenArr[i - 1];
                        group.AttValue["REDEND"] = 0;
                        group.AttValue["AMBER"] = 3;
                    }
                    else
                    {
                        group.AttValue["GREENEND"] = infoArr[i - 1] - 3;
                        group.AttValue["REDEND"] = infoArr[i - 1] - secondgreenArr[i - 1] - 3;
                        group.AttValue["AMBER"] = 3;
                    }
                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            columnClear();
            if (peishifinished || fenshiduanFinished||dingshiFinished)
            {
                fangzhen = true;
                Thread t3 = new Thread(columnAdd);
                t3.Start();
                InitialSystem();

                startFlag = true;

                for (itemTimeStep = 1; itemTimeStep <= vis.Simulation.Period; itemTimeStep++) //“循环单步仿真”主体
                {
                    System.Windows.Forms.Application.DoEvents(); //外部按键事件读取

                    if (startFlag)
                    {
                        vis.Simulation.RunSingleStep(); //启动单步仿真                       
                        //获取检测器数据
                        Thread t2 = new Thread(detectData);
                        t2.Start();
                        //更新周期
                        label15.Text = "仿真时长 :" + Convert.ToString(itemTimeStep);
                        progressBar1.Visible = true;
                        progressBar1.Value++;
                        simuOver();
                    }
                    else {
                        MessageBox.Show("vissim未正常启动，请重启！！！");
                    }
                }
                //  timer1.Start();

                //lukouDetectorSum();
                RestSystem();
            }
            else {
                MessageBox.Show("请输入路口信息，再点击仿真！！！");
            }
        }
        public void RestSystem() {
            button4.Text = "暂停";
            // vis.Exit();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            skinEngine1.SkinFile = Application.StartupPath + @"\Skins\EmeraldColor1.ssk";
            try
            {
                webBrowser1.Navigate(Application.StartupPath + @"\GoogleMap.htm");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "异常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }
        //导入路网文件

        private void textBox1_TextChanged(object sender, EventArgs e)
        {



        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void zhixiaokeche_TextChanged(object sender, EventArgs e)
        {


        }

        private void zhidakeche_TextChanged(object sender, EventArgs e)
        {


        }

        private void zhidahuohce_TextChanged(object sender, EventArgs e)
        {


        }

        private void youxiaokeche_TextChanged(object sender, EventArgs e)
        {


        }

        private void youdakeche_TextChanged(object sender, EventArgs e)
        {


        }

        private void youdahuoche_TextChanged(object sender, EventArgs e)
        {


        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void zhidahuoche_TextChanged(object sender, EventArgs e)
        {


        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {


        }

        private void youzhuanxiangweiliuliang_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //foreach (ListViewItem lst in listView1.SelectedItems) {
            //    lst.Remove();
            //}
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            startFlag = !startFlag;

            if (startFlag)
            {
                this.button4.Text = "暂停";
            }
            else
            {
                this.button4.Text = "继续";
            }

            while (!startFlag)
            {
                System.Windows.Forms.Application.DoEvents(); //外部按键事件读取
            }
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {


        }


        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView2_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }
        //DataCollection检测器的数据获取
        public int DataCollectionValue(int dataCollectionNo, string value) //DataCollection检测器函数 输入：dataCollectionNo dataCollection检测器编号， value 检测属性
        {

            DataCollection dataCollection = vis.Net.DataCollections.GetDataCollectionByNumber(dataCollectionNo);
            switch (value)
            {
                case "OCCUPANCYRATE":
                    return Convert.ToInt32(dataCollection.GetResult("OCCUPANCYRATE", "SUM", 0));
                case "SPEED":
                    return Convert.ToInt32(dataCollection.GetResult("SPEED", "MEAN", 0));
                case "NVEHICLES":
                    return Convert.ToInt32(dataCollection.GetResult("NVEHICLES", "SUM", 0));
                default: return 0;
            }
        }
        //排队长度检测参数的获取
        public int QueueCounterValue(int queueCounterNo, String value) //DataCollection检测器函数 输入：dataCollectionNo dataCollection检测器编号(路段编号)， value 检测属性
        {

            QueueCounter queueCounter = vis.Net.QueueCounters.GetQueueCounterByNumber(queueCounterNo);
            switch (value)
            {
                case "MEAN":
                    return Convert.ToInt32(queueCounter.GetResult(itemTimeStep, "MEAN"));
                case "STOPS":
                    return Convert.ToInt32(queueCounter.GetResult(itemTimeStep, "STOPS"));
                default:
                    return 0;
            }
        }

        //延误检测器的获取
        private int DelayEvaluationValue(int delayNum, String value) {
            Delay delay = vis.Net.Delays.GetDelayByNumber(delayNum);
            switch (value)
            {
                case "DELAY":
                    return Convert.ToInt32(delay.GetResult(itemTimeStep, "DELAY", "", 0));
                // MessageBox.Show(delay.GetResult(600, "DELAY", "", 0));
                //case "STOPPEDDELAY":
                //    return Convert.ToInt32(delay.GetResult(itemTimeStep, "NSTOPS","",0));
                default: return 0;
            }

        }
        public void detectData() {

            DataCollectionEvaluation dtDataCollectionEvaluation;
            dtDataCollectionEvaluation = vis.Evaluation.DataCollectionEvaluation;

            //dtDataCollection = vis.Net.DataCollections.GetDataCollectionByNumber(1);
            //int OccupancyRate = Convert.ToInt32(dtDataCollection.GetResult("OCCUPANCYRATE", "SUM", 0)); //获取占有率参数
            //int Speed = Convert.ToInt32(dtDataCollection.GetResult("SPPED", "MEAN", 0));
            //dataCollection数据获取

            if (itemTimeStep % detectPeriod == 0 && itemTimeStep != 0 && itemTimeStep <= detectTime)
            {
                for (int item_dataCol = 1; item_dataCol <= 4; item_dataCol++)
                {
                    dataCollectionSpeed[itemTimeStep / detectPeriod - 1, item_dataCol - 1] = DataCollectionValue(item_dataCol, "SPEED"); //记录所有DataCollections速度
                    dataCollectionOccupancyRate[itemTimeStep / detectPeriod - 1, item_dataCol - 1] = DataCollectionValue(item_dataCol, "OCCUPANCYRATE"); //记录所有DataCollections占有率
                    dataCollectionVehiclesNum[itemTimeStep / detectPeriod - 1, item_dataCol - 1] = DataCollectionValue(item_dataCol, "NVEHICLES"); //记录所有DataCollections过车数                                                                                                                                               //System.Threading.Thread.Sleep(30000);

                }
                for (int item_dataCol = 1; item_dataCol <= 4; item_dataCol++)
                {
                    queueCounterLength[itemTimeStep / detectPeriod - 1, item_dataCol - 1] = QueueCounterValue(item_dataCol, "MEAN");
                    queueCounterStops[itemTimeStep / detectPeriod - 1, item_dataCol - 1] = QueueCounterValue(item_dataCol, "STOPS");
                }

                for (int item_dataCol = 1; item_dataCol <= 4; item_dataCol++)
                {
                    delayDelay[itemTimeStep / detectPeriod - 1, item_dataCol - 1] = DelayEvaluationValue(item_dataCol, "DELAY");
                    //delayStopped[itemTimeStep / detectPeriod - 1, item_dataCol - 1] = DelayEvaluationValue(item_dataCol, "NSTOPS");
                }

            }

        }
        private void comboBox2_SelectedIndexChanged_1(object sender, EventArgs e)
        {


        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (fangzhen == true) {
                String dataCollectionIndex = comboBox3.SelectedItem.ToString(); //comboBoxDataCollectionNo 为下拉框控件的名称，dataCollectionNo 为路口编号

                //检测数据显示//

                dcqueue.Rows.Clear();

                for (int t = 1; t <= detectTime / detectPeriod; t++) //detectTimeSeriesNum 检测时刻个数（如: 仿真周期3600s，检测周期300s，则检测时刻个数为12）
                {
                    DataRow dr = dcqueue.NewRow();
                    dr["检测时间"] = t * detectPeriod; //detectPeriod;检测周期
                    dr["平均排队长度"] = queueCounterLength[t - 1, Convert.ToInt32(dataCollectionIndex) - 1]; //dataCollectionVehiclesNum[,] 检测参数存储数组
                    dr["停车次数"] = queueCounterStops[t - 1, Convert.ToInt32(dataCollectionIndex) - 1]; //dataCollectionSpeed[,] 同上               
                    dcqueue.Rows.Add(dr);
                }
                dataGridView3.DataSource = dcqueue; //显示数据表 this.dataGridViewDataCollection 为自定义的DataGridView控件名
                                                    //保存数据
                if (checkBox2.Checked)
                {
                    String excelFilePath = "C:\\Users\\han\\Desktop\\simuData";

                    String excelFileName = "_" + this.tabControl2.SelectedTab.Name + "_" + dataCollectionIndex; //保存数据的文件名

                    OutDataToExcel(dcqueue, excelFilePath, excelFileName); //写入Excel文件

                }
            }
        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView2_CellContentClick_2(object sender, DataGridViewCellEventArgs e)
        {

        }

        public static void OutDataToExcel(System.Data.DataTable srcDataTable, string excelFilePath, string excelFileName)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            object missing = System.Reflection.Missing.Value;

            //导出到execl   
            try
            {
                if (xlApp == null)
                {
                    MessageBox.Show("无法创建Excel对象，可能您的电脑未安装Excel!");

                    return;
                }

                Microsoft.Office.Interop.Excel.Workbooks xlBooks = xlApp.Workbooks;

                Microsoft.Office.Interop.Excel.Workbook xlBook = xlBooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);

                Microsoft.Office.Interop.Excel.Worksheet xlSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlBook.Worksheets[1];

                //让后台执行设置为不可见，为true的话会看到打开一个Excel，然后数据在往里写  
                xlApp.Visible = false;

                object[,] objData = new object[srcDataTable.Rows.Count + 1, srcDataTable.Columns.Count];

                //首先将数据写入到一个二维数组中  
                for (int i = 0; i < srcDataTable.Columns.Count; i++)
                {
                    objData[0, i] = srcDataTable.Columns[i].ColumnName;
                }
                if (srcDataTable.Rows.Count > 0)
                {
                    for (int i = 0; i < srcDataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < srcDataTable.Columns.Count; j++)
                        {
                            objData[i + 1, j] = srcDataTable.Rows[i][j];
                        }
                    }
                }

                string startCol = "A";

                int iCnt = (srcDataTable.Columns.Count / 26);

                string endColSignal = (iCnt == 0 ? "" : ((char)('A' + (iCnt - 1))).ToString());

                string endCol = endColSignal + ((char)('A' + srcDataTable.Columns.Count - iCnt * 26 - 1)).ToString();

                Microsoft.Office.Interop.Excel.Range range = xlSheet.get_Range(startCol + "1", endCol + (srcDataTable.Rows.Count - iCnt * 26 + 1).ToString());

                range.Value = objData; //给Exccel中的Range整体赋值  

                range.EntireColumn.AutoFit(); //设定Excel列宽度自适应  

                xlSheet.get_Range(startCol + "1", endCol + "1").Font.Bold = 1;//Excel文件列名 字体设定为Bold  

                //设置禁止弹出保存和覆盖的询问提示框  
                xlApp.DisplayAlerts = false;

                xlApp.AlertBeforeOverwriting = false;

                if (xlSheet != null)
                {
                    //xlSheet.SaveAs(excelFilePath, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                    if (excelFilePath.EndsWith("\\"))
                        //保存Excel文件
                        xlBook.SaveAs(excelFilePath + DateTime.Now.ToString("yyyMMddhhmm") + excelFileName + ".xlsx");
                    else
                        xlBook.SaveAs(excelFilePath + "\\" + DateTime.Now.ToString("yyyMMddhhmm") + excelFileName + ".xlsx");

                    MessageBox.Show("文件保存成功!");

                    System.Diagnostics.Process[] excelProcess = System.Diagnostics.Process.GetProcessesByName("EXCEL");

                    foreach (System.Diagnostics.Process p in excelProcess)
                    {
                        p.Kill();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("异常!");

                throw ex;
            }
        }
        private void comboBox2_SelectedIndexChanged_2(object sender, EventArgs e)
        {
            if (fangzhen == true) {//确保“启动仿真按钮首先点击”           
                String dataCollectionIndex = comboBox2.SelectedItem.ToString(); //comboBoxDataCollectionNo 为下拉框控件的名称，dataCollectionNo 为路口编号

                //检测数据显示//

                dtDataCollection.Rows.Clear();

                for (int t = 1; t <= detectTime / detectPeriod; t++) //detectTimeSeriesNum 检测时刻个数（如: 仿真周期3600s，检测周期300s，则检测时刻个数为12）
                {
                    DataRow dr = dtDataCollection.NewRow();
                    int lukouindex = Convert.ToInt32(dataCollectionIndex) - 1;

                    dr["检测时间"] = t * detectPeriod; //detectPeriod;检测周期
                    dr["车辆总数"] = dataCollectionVehiclesNum[t - 1, lukouindex] * 6; //dataCollectionVehiclesNum[,] 检测参数存储数组
                    dr["速度"] = dataCollectionSpeed[t - 1, lukouindex]; //dataCollectionSpeed[,] 同上
                    dr["占有率"] = dataCollectionOccupancyRate[t - 1, lukouindex]; //dataCollectionOccupancyRate 同上
                    dtDataCollection.Rows.Add(dr);
                }
                dataGridView2.DataSource = dtDataCollection; //显示数据表 this.dataGridViewDataCollection 为自定义的DataGridView控件名

                //保存数据
                if (checkBox1.Checked)
                {
                    String excelFilePath = "C:\\Users\\han\\Desktop\\simuData";
                    String excelFileName = "_" + this.tabControl2.SelectedTab.Name + "_" + dataCollectionIndex; //保存数据的文件名

                    //  Thread t2 = new Thread(new ParameterizedThreadStart(OutDataToExcel));

                    OutDataToExcel(dtDataCollection, excelFilePath, excelFileName); //写入Excel文件

                }

            }
        }

        private void groupBox7_Enter(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {


        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            //combobox4Arr[] = comboBox4.SelectedItem.ToString();
        }

        private void button6_Click(object sender, EventArgs e)
        {


        }
        public void fenshiduanShow() {
            if (String.IsNullOrEmpty(textBox1.Text) && String.IsNullOrEmpty(textBox3.Text) && String.IsNullOrEmpty(textBox3.Text))
            {
                MessageBox.Show("请输入定时控制完整信息！！！");
            }
            else {
                c0 = Convert.ToInt32(textBox1.Text);
                G = Convert.ToInt32(textBox3.Text);
                R = Convert.ToInt32(textBox2.Text);
                int xvalue = Convert.ToInt32(comboBox5.SelectedItem);
                infoArr[xvalue - 1] = c0;
                firstgreenArr[xvalue - 1] = G;
                secondgreenArr[xvalue - 1] = R;

                //waitingforSend = xvalue.ToString() + infoArr[xvalue - 1].ToString() + firstgreenArr[xvalue - 1].ToString() + secondgreenArr[xvalue - 1].ToString() + 3.ToString();
                waitingforSend = waitDataResult(xvalue, infoArr[xvalue - 1], firstgreenArr[xvalue - 1], secondgreenArr[xvalue - 1]);
                //if ((xvalue - 1) == 0&&(systemType==2))
                //{
                    listView3.View = View.Details;
                    listView3.FullRowSelect = true;
                    listView3.Columns.Add("路口编号", 100, HorizontalAlignment.Center);
                    listView3.Columns.Add("周期(s)", 100, HorizontalAlignment.Center);
                    listView3.Columns.Add("直行相位绿灯时间(s)", 100, HorizontalAlignment.Center);
                    listView3.Columns.Add("右转相位绿灯时间(s)", 100, HorizontalAlignment.Center);
                    listView3.Columns.Add("黄灯时间(s)", 100, HorizontalAlignment.Center);
              //  }
                ListViewItem itm = listView3.Items.Add(this.comboBox5.SelectedItem.ToString());
                itm.SubItems.AddRange(new string[] { c0.ToString(), G.ToString(), R.ToString(), 3.ToString() });
                for (int i = 0; i < listView3.Items.Count; i++)
                {
                    if (i % 2 == 0)
                    {
                        listView3.Items[i].BackColor = Color.Gray;
                    }
                    else
                    {
                        listView3.Items[i].BackColor = Color.LightGreen;
                    }
                }

                //柱状图生成
                chart6.Series[0].Points.AddXY(xvalue, c0);
                chart6.Series[1].Points.AddXY(xvalue, G);
                chart6.Series[2].Points.AddXY(xvalue, R);
                chart6.Series[3].Points.AddXY(xvalue, 3);
                fenshiduanFinished = true;
            }

        }
        public void dingshiShow() {
            if (String.IsNullOrEmpty(textBox4.Text) && String.IsNullOrEmpty(textBox5.Text) && String.IsNullOrEmpty(textBox6.Text))
            {
                MessageBox.Show("请输入定时控制完整信息！！！");
            }
            else
            {
                c0 = Convert.ToInt32(textBox4.Text);
                G = Convert.ToInt32(textBox5.Text);
                R = Convert.ToInt32(textBox6.Text);
                int xvalue = Convert.ToInt32(comboBox1.SelectedItem);
                infoArr[xvalue - 1] = c0;
                firstgreenArr[xvalue - 1] = G;
                secondgreenArr[xvalue - 1] = R;

               // waitingforSend = xvalue.ToString() + infoArr[xvalue - 1].ToString() + firstgreenArr[xvalue - 1].ToString() + secondgreenArr[xvalue - 1].ToString() + 3.ToString();
                waitingforSend = waitDataResult(xvalue, infoArr[xvalue - 1], firstgreenArr[xvalue - 1], secondgreenArr[xvalue - 1]);
                //if ((xvalue - 1) == 0&&systemType==3)
                //{
                listView4.View = View.Details;
                    listView4.FullRowSelect = true;
                    listView4.Columns.Add("路口编号", 100, HorizontalAlignment.Center);
                    listView4.Columns.Add("周期(s)", 100, HorizontalAlignment.Center);
                    listView4.Columns.Add("直行相位绿灯时间(s)", 100, HorizontalAlignment.Center);
                    listView4.Columns.Add("右转相位绿灯时间(s)", 100, HorizontalAlignment.Center);
                    listView4.Columns.Add("黄灯时间(s)", 100, HorizontalAlignment.Center);
                //}
                ListViewItem itm = listView4.Items.Add(this.comboBox1.SelectedItem.ToString());
                itm.SubItems.AddRange(new string[] { c0.ToString(), G.ToString(), R.ToString(), 3.ToString() });
                for (int i = 0; i < listView4.Items.Count; i++)
                {
                    if (i % 2 == 0)
                    {
                        listView4.Items[i].BackColor = Color.Gray;
                    }
                    else
                    {
                        listView4.Items[i].BackColor = Color.LightGreen;
                    }
                }

                //柱状图生成
                chart7.Series[0].Points.AddXY(xvalue, c0);
                chart7.Series[1].Points.AddXY(xvalue, G);
                chart7.Series[2].Points.AddXY(xvalue, R);
                chart7.Series[3].Points.AddXY(xvalue, 3);
                dingshiFinished = true;
            }

        }

        public void peishiInfoShow() {
            //输入项获取数据
            if (String.IsNullOrEmpty(zhixiaokeche.Text)&& String.IsNullOrEmpty(zhidakeche.Text) && String.IsNullOrEmpty(zhixiaohuoche.Text) && String.IsNullOrEmpty(zhidahuoche.Text) && String.IsNullOrEmpty(youxiaokeche.Text) && String.IsNullOrEmpty(youdakeche.Text) && String.IsNullOrEmpty(youxiaohuoche.Text) && String.IsNullOrEmpty(youdahuoche.Text)) {
                MessageBox.Show("请输入路口完整交通信息！！！");
            }
            else { 
            int aa1 = int.Parse(zhixiaokeche.Text);
            int bb1 = int.Parse(zhidakeche.Text);
            int cc1 = int.Parse(zhixiaohuoche.Text);
            int dd1 = int.Parse(zhidahuoche.Text);
            int aa2 = int.Parse(youxiaokeche.Text);
            int bb2 = int.Parse(youdakeche.Text);
            int cc2 = int.Parse(youxiaohuoche.Text);
            int dd2 = int.Parse(youdahuoche.Text);

            label10.Visible = true;
            label9.Visible = true;
            label10.Text = "";
            label9.Text = "";
            //计算直行和右转相位算出的流量并显示
            qd1 = (int)(((aa1 + 1.5 * bb1 + cc1 + 1.5 * dd1) / (180 * 4)) * 3600);
            qd2 = (int)(((aa2 + 1.5 * bb2 + cc2 + 1.5 * dd2) / (180 * 2)) * 3600);
            sd = int.Parse(zongbaoheliuliang.Text);
            double sd1 = (double)sd;

            label10.Text = "直行相位流量：" + qd1.ToString();
            label9.Text = "右转相位流量：" + qd2.ToString();

            double Y = (qd1 + qd2) / sd1;
            //计算周期，存数组
            int c0 = (int)(20.0 / (1.0 - Y));

            if (c0 < 0)
            {
                MessageBox.Show("结果为负，输入信息有误，请重新输入！");
            }

            int d = qd1 + qd2;
            double D = (double)d;
            //计算每个相位的灯色时间，存数组
            G = (int)((c0 - 10) * qd1 / D + 2);
            R = (int)((c0 - 10) * qd2 / D + 2);
            int xvalue = Convert.ToInt32(comboBox4.SelectedItem);
            //存数组操作
            infoArr[xvalue - 1] = c0;
            firstgreenArr[xvalue - 1] = G; 
            secondgreenArr[xvalue - 1] = R;

            //waitingforSend = xvalue.ToString() + infoArr[xvalue - 1].ToString() + firstgreenArr[xvalue - 1].ToString() + secondgreenArr[xvalue - 1].ToString()+3.ToString();
             waitingforSend = waitDataResult(xvalue, infoArr[xvalue - 1], firstgreenArr[xvalue - 1], secondgreenArr[xvalue - 1]);
             
        
                
                //  if ((xvalue - 1) == 0)
          // {
                listView1.View = View.Details;
                listView1.FullRowSelect = true;
                listView1.Columns.Add("路口编号", 100, HorizontalAlignment.Center);
                listView1.Columns.Add("周期(s)", 100, HorizontalAlignment.Center);
                listView1.Columns.Add("直行相位绿灯时间(s)", 100, HorizontalAlignment.Center);
                listView1.Columns.Add("右转相位绿灯时间(s)", 100, HorizontalAlignment.Center);
                listView1.Columns.Add("黄灯时间(s)", 100, HorizontalAlignment.Center);
          //  }
            ListViewItem itm = listView1.Items.Add(this.comboBox4.SelectedItem.ToString());
            itm.SubItems.AddRange(new string[] { c0.ToString(), G.ToString(), R.ToString(), 3.ToString() });
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                if (i % 2 == 0)
                {
                    listView1.Items[i].BackColor = Color.Gray;
                }
                else
                {
                    listView1.Items[i].BackColor = Color.LightGreen;
                }
            }
  
            //柱状图生成
            chart1.Series[0].Points.AddXY(xvalue, c0);
            chart1.Series[1].Points.AddXY(xvalue, G);
            chart1.Series[2].Points.AddXY(xvalue, R);
            chart1.Series[3].Points.AddXY(xvalue, 3);
                peishifinished = true;
            }
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            
            systemType = 1;
            showControll();
            if (systemType == 1) {
                peishiInfoShow();
                send();
            }
                   
        }
 
        public void showSpeedZhexiantu() {
            
            for (int t = 1; t <= detectTime/detectPeriod; t++)
            {
                chart2.Series[0].Points.AddXY(t * detectPeriod, dataCollectionSpeed[t - 1,0]);
                chart2.Series[1].Points.AddXY(t * detectPeriod, dataCollectionSpeed[t - 1,1]);
                chart2.Series[2].Points.AddXY(t * detectPeriod, dataCollectionSpeed[t - 1,2]);
                chart2.Series[3].Points.AddXY(t * detectPeriod, dataCollectionSpeed[t - 1,3]);
            }
        }
        public void showStopsZhexiantu()
        {

            for (int t = 1; t <= detectTime/detectPeriod; t++)
            {
                chart5.Series[0].Points.AddXY(t * detectPeriod, queueCounterStops[t - 1, 0]);
                chart5.Series[1].Points.AddXY(t * detectPeriod, queueCounterStops[t - 1, 1]);
                chart5.Series[2].Points.AddXY(t * detectPeriod, queueCounterStops[t - 1, 2]);
                chart5.Series[3].Points.AddXY(t * detectPeriod, queueCounterStops[t - 1, 3]);
            }
        }
        public void showLengthZhexiantu()
        {

            for (int t = 1; t <= detectTime/detectPeriod; t++)
            {
                chart4.Series[0].Points.AddXY(t * detectPeriod, queueCounterLength[t - 1, 0]);
                chart4.Series[1].Points.AddXY(t * detectPeriod, queueCounterLength[t - 1, 1]);
                chart4.Series[2].Points.AddXY(t * detectPeriod, queueCounterLength[t - 1, 2]);
                chart4.Series[3].Points.AddXY(t * detectPeriod, queueCounterLength[t - 1, 3]);
            }
        }


        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void tabPage8_Click(object sender, EventArgs e)
        {

        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            //检测数据显示//
            if (fangzhen == true) { 
            String dataCollectionIndex = comboBox6.SelectedItem.ToString(); //comboBoxDataCollectionNo 为下拉框控件的名称，dataCollectionNo 为路口编号
            delayData.Rows.Clear();

            for (int t = 1; t <= detectTime/detectPeriod; t++) //detectTimeSeriesNum 检测时刻个数（如: 仿真周期3600s，检测周期300s，则检测时刻个数为12）
            {
                DataRow dr = delayData.NewRow();
                int lukouindex = Convert.ToInt32(dataCollectionIndex) - 1;

                dr["检测时间"] = t * detectPeriod; //detectPeriod;检测周期
                dr["单车平均延误"] = delayDelay[t - 1, lukouindex]; //dataCollectionVehiclesNum[,] 检测参数存储数组
              
                delayData.Rows.Add(dr);
            }
            dataGridView4.DataSource = delayData; //显示数据表 this.dataGridViewDataCollection 为自定义的DataGridV
                if (checkBox3.Checked) {
                    //String excelFilePath = showSaveFileDialog();
                     String excelFilePath = "C:\\Users\\han\\Desktop\\simuData";
                   // String excelFileName = excelFilePath.Substring(excelFilePath.LastIndexOf("\\") + 1);
                     String excelFileName = "_" + this.tabControl2.SelectedTab.Name + "_" + dataCollectionIndex; //保存数据的文件名

                    OutDataToExcel(delayData, excelFilePath, excelFileName); //写入Excel文件
                }
          }
        }
        //保存数据到特定的路径+


        private string showSaveFileDialog() {
            string localFilePath = "";
            //string localFilePath, fileNameExt, newFileName, FilePath; 
            SaveFileDialog sfd = new SaveFileDialog();
            //设置文件类型 
            sfd.Filter = "Excel表格（*.xls）|*.xls";

            //设置默认文件类型显示顺序 
            sfd.FilterIndex = 1;

            //保存对话框是否记忆上次打开的目录 
            sfd.RestoreDirectory = true;

            //点了保存按钮进入 
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                localFilePath = sfd.FileName.ToString(); //获得文件路径 
                string fileNameExt = localFilePath.Substring(localFilePath.LastIndexOf("\\") + 1); //获取文件名，不带路径

                //获取文件路径，不带文件名 
                //FilePath = localFilePath.Substring(0, localFilePath.LastIndexOf("\\")); 

                //给文件名前加上时间 
                //newFileName = DateTime.Now.ToString("yyyyMMdd") + fileNameExt; 

                //在文件名里加字符 
                //saveFileDialog1.FileName.Insert(1,"dameng"); 

                //System.IO.FileStream fs = (System.IO.FileStream)sfd.OpenFile();//输出文件 

                ////fs输出带文字或图片的文件，就看需求了 
            }

            return localFilePath;

        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Now;         
            label21.Text = dt.ToString();
        }
        private void listView2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            
            showSpeedZhexiantu();
            showLengthZhexiantu();
            showStopsZhexiantu();
            showDelayZhexiantu();
           
        }
        private void showDelayZhexiantu() {
            for (int t = 1; t <= detectTime/detectPeriod; t++)
            {
                chart3.Series[0].Points.AddXY(t * detectPeriod, delayDelay[t - 1, 0]);
                chart3.Series[1].Points.AddXY(t * detectPeriod, delayDelay[t - 1, 1]);
                chart3.Series[2].Points.AddXY(t * detectPeriod, delayDelay[t - 1, 2]);
                chart3.Series[3].Points.AddXY(t * detectPeriod, delayDelay[t - 1, 3]);
            }

        }
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chart3_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            label29.Visible = true;
            systemType = 2;
            showControll();
            if (systemType == 2) {
                fenshiduanShow();
                send();
            }
            label29.Text = numericUpDown1.Value.ToString() + "时" + numericUpDown3.Value.ToString() + "分" + "-----" + numericUpDown2.Value.ToString() + "时" + numericUpDown4.Value.ToString() + "分";

        }

        private void listView3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_2(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void chart6_Click(object sender, EventArgs e)
        {

        }

        private void label26_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {

        }

        private void groupBox7_Enter_1(object sender, EventArgs e)
        {

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {

        }
        private void showControll() {
            label31.Visible = true;
            switch (systemType) {
                case 1:
                    label31.Text = "多智能体控制";
                    break;
                case 2:
                    label31.Text = "分时段定时控制";
                    break;
                case 3:
                    label31.Text = "定时控制";
                    break;
                default:
                    label31.Text = "多智能体控制";
                    break;
            }
        }
        //仿真结束时，红绿灯信息清空，等待下一次或者下一个控制策略存入
        private void simuOver() {
            if (itemTimeStep == vis.Simulation.Period) {
                Array.Clear(infoArr,0,infoArr.Length);
                Array.Clear(firstgreenArr, 0, firstgreenArr.Length);
                Array.Clear(secondgreenArr, 0, secondgreenArr.Length);
                progressBar1.Value = 0;
            }
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            webBrowser1.Document.InvokeScript("setLocation", new object[] { 106.5550451816, 29.5588616901 });
            
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            
        }

        private void skinButton1_Click(object sender, EventArgs e)
        {

        }

        private void label35_Click(object sender, EventArgs e)
        {

        }

        private void chart7_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            systemType = 3;
            showControll();
            if (systemType == 3)
            {
                dingshiShow();
                send();
            }
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void listView4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {

        }
        //UDP发送数据
        private void send() {
            UdpClient udpClient = new UdpClient(10001);

            udpClient.Connect(IPAddress.Parse("172.20.33.145"), 1002);

            Byte[] sendBytes = Encoding.Default.GetBytes(waitingforSend);
            Console.WriteLine("已发送数据：" + waitingforSend);

            udpClient.Send(sendBytes, sendBytes.Length);

            udpClient.Close();

        }

        private void groupBox9_Enter(object sender, EventArgs e)
        {

        }

        private void label31_Click(object sender, EventArgs e)
        {

        }
        private String waitDataResult(int s1, int s2, int s3,int s4) {
            
            String packResult = "路口ID" + s1.ToString() + "周期" + s2.ToString() + "绿灯" + s3.ToString() + "红灯" + s4.ToString() + "黄灯" + 3.ToString();
            
            return packResult;
        }
        //excel数据处理

        private int ExcelHandling() {
            return 0;
        }
        
    }
}

