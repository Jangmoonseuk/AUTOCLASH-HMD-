using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Threading;
using System.Reflection;

using SHConnector;

using autoCLASH.Lib;

using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace autoCLASH
{
    public partial class ClashControl : UserControl, IEntryConnector
    {
        private ListViewItem lvItem;

        private IVIZZARDService Connector;
        private List<ClashTaskMultiVO> TaskList { get; set; }
        private List<ClashTaskVO> TaskListWire { get; set; }

        public float SHIP_MIN_Z { get; set; }
        public float SHIP_MAX_Z { get; set; }

        public string CLASH_MODEL_FILE { get; set; }

        //
        string _computerName = Environment.GetEnvironmentVariable("COMPUTERNAME");
        string initpath = @"Q:\env\Vitesse\Infoget\autoAP\initFile.txt";
        string _tempDir = @"C:\Temp"; //Environment.GetEnvironmentVariable("TEMP");
        string _tempDir2 = "";//밑에서 재 정의.

        string am12EXEDir = @"C:\Aveva\Marine\OH12.1.SP4";
        string proj = "";
        string am12MDB = "/HULL";
        string logFileName = "";
        bool marinbool = false;
        bool dirbool = false;
        string filenameM = "";
        string filepathM = "";
        bool lastSTEP = true;
        bool lugwire = false;

        string am12USERNAME = "SYSTEM";
        string am12PASSWORD = "XXXXXX";

        string _userName = Environment.GetEnvironmentVariable("USERNAME");
        string _savedUser = "";
        List<string> exceptBlock = new List<string>();

        Process myProc = Process.GetCurrentProcess();

        int columnNo = 3;
        double totalLoading = 0;
        double totalcheck = 0;
        double totalend = 0;
        bool autorun = false;
        bool allblock = false;

        string allblockpath = "";
        string companyName = "IFG"; //밑에서 재 정의.
        string _originPath = @"D:\autoCCS\Ship"; //밑에서 재 정의

        string _vizzardDir = "";
        string LocalPath = @"C:\VIZZARD64_Design\Plugins";

        string _shipNo = ""; //030737970
        string logfilePath = "";//로그파일 path

        string _selectedSHIPPath = ""; // D:\ZAUTO_AP\030737970
        string _erectPath ="";
        string _excelFile = ""; //밑에서 재 정의.
        string selectedNodeLast = "";
        DateTime starttime = DateTime.Now;
        TimeSpan TotalLoad = new TimeSpan();

        int maxCols = 9; //

        private int _nIndex = 0;
        private bool _checkOver = false;

        Dictionary<string, ObjectPropertyVO> dic_node2QVO = new Dictionary<string, ObjectPropertyVO>();

        Dictionary<string, string> dic_treenode2Path = new Dictionary<string, string>();
        string _treeAllFile = @"";

        Dictionary<string, string> dic_Block2Date = new Dictionary<string, string>();
        string _erectDateFile = @"";

        Dictionary<string, string> dic_Block2No = new Dictionary<string, string>();
        Dictionary<string, string> dic_BlockDir = new Dictionary<string, string>();
        Dictionary<string, string> dic_BlockLength = new Dictionary<string, string>();
        List<string> runcomputer = new List<string>();
        List<string> erectBlocks = new List<string>();//탑재블록 넣어두기
        string _erectSavedFile = @""; //사용자 탑재순서 결정
        bool worker = false;

        string _movingName = "";
        float _movingValue = 0;

        float _minHeight = 999999;
        float _maxHeight = -999999;

        int _loopMAX = 10;

        decimal _allowTimeMIN = 20190101235959.99m;
        decimal _allowTimeMAX = 20191231235959.99m;

        ArrayList _selectedItems = new ArrayList();

        double _dMemory = 0;

        int _iIndex = 0;

        ArrayList _usedCSVFileNames = new ArrayList();
        string _csvFileName = "";
        string _csv2csvEXE = Path.Combine(@"C:\Aveva\Marine\OH12.1.SP4", @"csv2csvIFG.exe");

        List<string> gijangblock = new List<string>();

        string s_attFileNames = "";
        //

        public ClashControl()
        {
            InitializeComponent();
        }

        public ClashControl(IVIZZARDService conn) : this()
        {
            try
            {
                Connector = conn;

                Console.WriteLine("Start, ClashControl");

                if (Directory.Exists(@"M:\HMD") == true)
                {
                    companyName = "HMD";
                    _originPath = @"\\210.118.131.6\simulation\__Simulation_Program_Server\AMPROJ_REV";
                    //_originPath = @"C:\temp";

                    am12USERNAME = "D337935";
                    am12PASSWORD = "HMD337935";

                    _allowTimeMAX = 20991231235959.99m;
                }
                if(Environment.MachineName=="DESKTOP-MIR8VRM"|| Environment.MachineName == "DESKTOP-SERVER2")
                {
                    _allowTimeMAX = 20991231235959.99m;
                }
                _excelFile = Path.Combine(_originPath, "screen.xls");
                Console.WriteLine("companyName:{0}, _originPath:{1}, _excelFile:{2}", companyName, _originPath, _excelFile);
                if(companyName=="IFG")
                {
                    initpath = @"D:\TEST.txt";
                }
                //License Check
                bool bLicenseOK = LicenseCheck();
                if (bLicenseOK == false) { MessageBox.Show("License Expired..."); return; }

                Console.WriteLine("Before, InitDrive");
                InitDrive();
                Console.WriteLine("After, InitDrive");

                /*
                this.watListView1.ListViewItemSorter = new ListViewItemComparer(3, "asc");
                watListView1.Sorting = SortOrder.Ascending;
                watListView1.Sort();
                watListView1.Refresh();
                */

                Console.WriteLine("~1");
                //

                //Connector.OnFinishedClashTestEvent += Connector_OnFinishedClashTestEvent;

                string strExcludingRules = @"C:\Temp";
                Connector.LoadClashTestExcludingRules(strExcludingRules);
                //
                Console.WriteLine("~2");

                //자동실행 이벤트
                try
                {
                    string[] Args = Connector.GetApplicationArguments();

                    arg = Args.ToList();
                    //arg.Add("/PLUGIN:\"C:\\Program Files\\Softhills\\VIZZARD Manager\\V3.0.2.19384\\Addins\\autoCLASH.xml\"");
                    //arg.Add("RUN");
                    //arg.Add("9");
                    //arg.Add("C:\\Temp\\267606\\B11C.REV");
                    //Args = arg.ToArray();

                    if (Args.Count() >= 3 && Args[2] == "1")
                    {
                        Connector.OnInitializedAppEvent += Connector_OnInitializedAppEvent;//정적 동적 간섭체크 by 자동수행기
                    }
                    if (Args.Count() >= 3 && Args[2] == "2")
                    {
                        Connector.OnInitializedAppEvent += Connector_OnInitializedAppEvent2;//Lug & Wire 간섭체크 by 자동수행기
                    }
                    if (Args.Count() >= 3 && Args[2] == "3")
                    {
                        Connector.OnInitializedAppEvent += Connector_OnInitializedAppEvent3;//정적 동적 간섭체크 all Block by command Line
                    }
                    if (Args.Count() >= 3 && Args[2] == "8")
                    {
                        Connector.OnInitializedAppEvent += Connector_OnInitializedAppEvent4;//정적 동적 간섭체크 all Block by command Line
                    }
                    if (Args.Count() >= 3 && Args[2] == "9")
                    {
                        Connector.OnInitializedAppEvent += Connector_OnInitializedAppEvent5;//정적 동적 간섭체크 all Block by command Line
                    }
                }
                catch { }
                try
                {
                    StreamReader init = new StreamReader(initpath);
                    List<string> readdata = new List<string>();
                    string line;
                    while ((line = init.ReadLine()) != null)
                    {
                        readdata.Add(line);
                    }
                    init.Close();
                    foreach (var a in readdata)
                    {
                        if (a.Contains('!'))
                            continue;
                        if (a.Contains("RangeTolerance="))
                        {
                            textBoxNear.Text = a.Split('=').Last();
                        }
                        if (a.Contains("ContractTolerance"))
                        {
                            textBoxContact.Text = a.Split('=').Last();
                        }
                        if (a.Contains("LBD_RUN"))
                        {
                            runcomputer.AddRange(a.Split('=').Last().Split(',').ToList());
                        }
                        if (a.Contains("RangeTolerance("))
                        {
                            gijang.Text = a.Split('=').Last();
                            gijangblock = a.Split('=').First().Replace("RangeTolerance", "").Replace("(", "").Replace(")", "").Split(',').ToList();
                        }
                        if (a.Contains("SKIP_LAST_STEP"))
                        {
                            if (a.Split('=').Last() == "0")
                                lastSTEP = true;
                            if (a.Split('=').Last() == "1")
                                lastSTEP = false;
                        }
                        if (a.Contains("ERECT_SKIP_BLOCK"))
                        {
                            exceptBlock.AddRange(a.Split('=').Last().Split(',').ToList());
                        }
                        if (a.Contains("INCLUDE_LUG_WIRE"))
                        {
                            if (a.Split('=').Last() == "0")
                                lugwire = false;
                            if (a.Split('=').Last() == "1")
                                lugwire = true;
                        }
                    }
                }
                catch
                { }

                string pcName = Environment.MachineName;
                if(pcName.Length<6)
                {
                    pcName = "AAAAAAAAA";
                }
                pcName = pcName.Substring(pcName.Length - 5, 5);
                if (!runcomputer.Contains(pcName))
                {
                    //buttonSaveErectNo.Enabled = false;//탑재순서 저장
                    button1.Enabled = false;//간섭체크 실행 막기
                    button2.Enabled = false;//Lug&Wire 실행 막기
                    worker = true;
                }
                if (companyName == "IFG")
                {
                    buttonSaveErectNo.Enabled = true;//탑재순서 저장
                    button1.Enabled = true;//간섭체크 실행 막기
                    button2.Enabled = true;//Lug&Wire 실행 막기
                    worker = false;
                    exceptBlock.Add("2B11"); exceptBlock.Add("2E12");
                }
            }
            catch(Exception exes) { MessageBox.Show(exes.Message); }
        }
        List<string> arg = new List<string>();
        List<string> selectblock = new List<string>();
        private void Connector_OnInitializedAppEvent(object sender, EventArgs e)
        {
            //프로그램이 구동 완료 됨.
            //자동 실행 프로그램
            autorun = true;
            cbBox1.Text = arg[1];
            Connector.RefreshReviewControl();
            //List<string> BlockList = new List<string>();
            //if (arg.Count >= 4)
            //{
            //    BlockList = arg[3].Split(',').ToList();
            //}
            logFileName = arg[3];//아큐먼트 add

            double work = double.Parse(arg[4]);//현재 작업
            double Allwork = double.Parse(arg[5]);//전체 작업

            this.watListView1.ListViewItemSorter = new ListViewItemComparer(6, "asc");
            watListView1.Sort();

            List<PluginVO> Plugins = Connector.GetPluginControlList();
            foreach (PluginVO plugin in Plugins)
            {
                // Plugin.xml Menu Title Attribute 정의한 텍스트 
                if (plugin.MenuTitle == "11.간섭체크(정적/동적/러그_와이어)")
                {
                    ((DevExpress.XtraBars.BarCheckItem)plugin.MenuButtonControl).PerformClick();
                }
            }

            cbBox1.Text = arg[1];
            cbBox1_SelectedIndexChanged(null, null);
            watListView1.BeginUpdate();
            int count = 1;
            foreach(ListViewItem item in watListView1.Items)
            {
                if (watListView1.Items.Count / Allwork * (work - 1) < count && count <= watListView1.Items.Count / Allwork * (work))
                {
                    item.Selected = true;
                    selectblock.Add(item.SubItems[0].Text);
                }
                count++;
            }
            watListView1.EndUpdate();
            autorun = true;

            button1_Click(null, null);
        }
        private void Connector_OnInitializedAppEvent2(object sender, EventArgs e)
        {
            //프로그램이 구동 완료 됨.
            //자동 실행 프로그램
            autorun = true;
            cbBox1.Text = arg[1];
            List<PluginVO> Plugins = Connector.GetPluginControlList();
            foreach (PluginVO plugin in Plugins)
            {
                // Plugin.xml Menu Title Attribute 정의한 텍스트 
                if (plugin.MenuTitle == "11.간섭체크(정적/동적/러그_와이어)")
                {
                    ((DevExpress.XtraBars.BarCheckItem)plugin.MenuButtonControl).PerformClick();
                }
            }
            tabControl1.SelectedIndex = 3;
            listView2.BeginUpdate();
            foreach (ListViewItem item in listView2.Items)
            {
                item.Selected = true;
            }
            listView2.EndUpdate();
            button2_Click(null, null);
        }
        private void Connector_OnInitializedAppEvent3(object sender, EventArgs e)
        {
            //프로그램이 구동 완료 됨.
            //자동 실행 프로그램
            autorun = true;
            marinbool = true;
            Stand.Checked = true;
            cbBox1.Text = arg[1];
            button3_Click(null, null);
            List<string> BlockList = new List<string>();
            List<string> tempsave = new List<string>();
            List<string> tempsave2 = new List<string>();

            if (arg.Count == 9)
            {
                filenameM = string.Format("{0}_{1}_{2}_{3}_{4}", arg[1], arg[7], arg[6], arg[5], arg[4]);
                filepathM = arg[8];
            }
            else
            { return; }
            //
            //
            if (arg.Count >= 4)
            {
                tempsave = arg[3].Split(',').ToList();
            }
            //
            _treeAllFile = Path.Combine(_selectedSHIPPath, @"Z99_TREE_ALL.TXT");
            if (File.Exists(_treeAllFile) == true)
            {
                string[] readLines = File.ReadAllLines(_treeAllFile, Encoding.Default);
                tree = readLines.ToList();
            }
            if(arg.Count>=5)
            {
                logFileName = arg[4];
            }

            foreach(var a in tempsave)
            {
                if(tree.Contains(a))
                {
                    tempsave2.Add(a);
                }
                else
                {
                    MessageBox.Show(string.Format("{0}의 경로가 잘못 되었습니다.", a));
                    return;
                }
            }
            //
            foreach (var a in tempsave2)
            {
                BlockList.Add(a.Split('/').Last());
            }
            List<PluginVO> Plugins = Connector.GetPluginControlList();
            foreach (PluginVO plugin in Plugins)
            {
                // Plugin.xml Menu Title Attribute 정의한 텍스트 
                if (plugin.MenuTitle == "11.간섭체크(정적/동적/러그_와이어)")
                {
                    ((DevExpress.XtraBars.BarCheckItem)plugin.MenuButtonControl).PerformClick();
                }
            }


            watListView1.BeginUpdate();
            int count = 0;
            foreach (ListViewItem item in watListView1.Items)
            {
                //if(count>=1)
                //{
                //    continue;
                //}
                if (arg.Count >= 4)
                {
                    if (!BlockList.Contains(item.SubItems[0].Text))
                    {
                        continue;
                    }
                }
                item.Selected = true;
                count++;
            }
            watListView1.EndUpdate();

            button1_Click(null, null);
        }
        private void Connector_OnInitializedAppEvent4(object sender, EventArgs e)
        {
            try
            {
                Console.WriteLine("Start, button1_Click");
                //로그파일생성
                string time2 = string.Format("{0}시{1}분{2}초", DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00")
        , DateTime.Now.Second.ToString("00"));
                string path2 = @"C:\temp";
                logfilePath = Path.Combine(path2, string.Format("{0}({1}_정동적실행).log", DateTime.Today.ToShortDateString(), time2));
                if (logFileName != "")
                {
                    logfilePath = logFileName;//arg에서 받은 Log파일 경로 설정
                }

                if (!(new DirectoryInfo(path2)).Exists)
                {
                    Console.WriteLine(string.Format("LOG 경로를 확인하세요 : {0}", path2));
                }
                Console.WriteLine("Log파일 이름 부여");

                StreamWriter log = new StreamWriter(logfilePath, true);
                starttime = DateTime.Now;
                log.WriteLine(string.Format("[{0} {1}]Log File", DateTime.Now.ToShortDateString(), time2));
                //log.WriteLine(string.Format("{0}_{1} 실행", DateTime.Now.ToShortDateString(),time));
                log.Close();
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
                return;
            }

            //프로그램이 구동 완료 됨.
            //자동 실행 프로그램
            autorun = true;
            marinbool = true;
            dirbool = true;
            filepathM = arg[3].Split(',').ToList()[0].Substring(0, arg[3].Split(',').ToList()[0].LastIndexOf('\\'));
            List<string> BlockList = new List<string>();

            //if (arg.Count == 9)
            //{
            //    filenameM = string.Format("{0}_{1}_{2}_{3}_{4}", arg[1], arg[7], arg[6], arg[5], arg[4]);
            //    filepathM = arg[8];
            //}
            //else
            //{ return; }
            foreach (var a in arg[3].Split(',').ToList())
            {
                BlockList.Add(a);
            }
            List<PluginVO> Plugins = Connector.GetPluginControlList();
            foreach (PluginVO plugin in Plugins)
            {
                // Plugin.xml Menu Title Attribute 정의한 텍스트 
                if (plugin.MenuTitle == "11.간섭체크(정적/동적/러그_와이어)")
                {
                    ((DevExpress.XtraBars.BarCheckItem)plugin.MenuButtonControl).PerformClick();
                }
            }
            ListViewItem bb = new ListViewItem();
            List<string> blocks = new List<string>();
            foreach (var a in arg[3].Split(',').ToList())
            {
                bb.SubItems[0].Text = a;
                blocks.Add(a);
            }
            watListView1.Items.Add(bb);
            Connector.AddDocuments(blocks.ToArray());

            ClashTaskVO taskModel = new ClashTaskVO();
            TaskList = new List<ClashTaskMultiVO>();
            // 간섭검사 기준 데이터 설정
            taskModel.CtType = CtType.CtTypeSelf; //정적간섭검사(모델) ->추가 2018.12.10.

            // 그룹 설정 : False (간섭검사 UI와 분리) / True (간섭검사 UI와 연동)
            taskModel.bInitTranslationValue = true;

            // 보이는 모델만
            taskModel.VisibleOnly = true;

            // 간섭검사 환경 설정
            // 어셈블리 단위 설정

            taskModel.CheckFixedGroupAsm = true;
            taskModel.CheckMovingGroupAsm = true;

            // 근접허용범위
            taskModel.bClashToleranceRange = false;
            taskModel.bRange = false;
            taskModel.RangeTolerance = Convert.ToInt32(textBoxNear.Text); //2.0f;

            // 접촉허용오차
            taskModel.bClashCalibrationTolerance = true;
            taskModel.bPenet = true;
            taskModel.ContractTolerance = Convert.ToInt32(textBoxContact.Text); //1.0f;               

            // 간섭제외 끝레벨
            taskModel.BottomLevel = int.Parse(textBoxExceptLevel.Text); //2;

            Console.WriteLine(" ~@3");

            List<NodeVO> items = Connector.GetChildObjects(0, ChildrenTypes.Children);
            for (int i = 0; i < items.Count; i++)
            {
                // 간섭검사 시작
                //Connector.Clash_StartCheck(task);
                string BLOCK_B = items[i].NodeName;
                Dictionary<string, string> fixedBlockDic = new Dictionary<string, string>();
                fixedBlockDic[BLOCK_B] = "";

                // 별도 쓰레드 처리를 위한 작업 목록에 작업 생성 및 추가
                ClashTaskMultiVO multiVOModel = new ClashTaskMultiVO();
                multiVOModel.Connector = Connector;
                multiVOModel.TaskVO = taskModel;
                multiVOModel.ListItem = null; // lviGroup;
                multiVOModel.MOVING_BLOCK = items[i].NodeName;
                multiVOModel.FIXED_BLOCK = fixedBlockDic.Keys.ToList();
                TaskList.Add(multiVOModel);
            }
            btnStartClashTest_Click(null, null);
        }
        private void Connector_OnInitializedAppEvent5(object sender, EventArgs e)
        {
            try
            {
                Console.WriteLine("Start, button1_Click");
                //로그파일생성
                string time2 = string.Format("{0}시{1}분{2}초", DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00")
        , DateTime.Now.Second.ToString("00"));
                string path2 = @"C:\temp";
                logfilePath = Path.Combine(path2, string.Format("{0}({1}_정동적실행).log", DateTime.Today.ToShortDateString(), time2));
                if (logFileName != "")
                {
                    logfilePath = logFileName;//arg에서 받은 Log파일 경로 설정
                }

                if (!(new DirectoryInfo(path2)).Exists)
                {
                    Console.WriteLine(string.Format("LOG 경로를 확인하세요 : {0}", path2));
                }
                Console.WriteLine("Log파일 이름 부여");

                StreamWriter log = new StreamWriter(logfilePath, true);
                starttime = DateTime.Now;
                log.WriteLine(string.Format("[{0} {1}]Log File", DateTime.Now.ToShortDateString(), time2));
                //log.WriteLine(string.Format("{0}_{1} 실행", DateTime.Now.ToShortDateString(),time));
                log.Close();
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
                return;
            }

            try
            {
                //프로그램이 구동 완료 됨.
                //자동 실행 프로그램
                //autorun = true;
                marinbool = true;
                dirbool = true;
                filepathM = arg[3].Split(',').ToList()[0].Substring(0, arg[3].Split(',').ToList()[0].LastIndexOf('\\'));
                List<string> BlockList = new List<string>();

                //if (arg.Count == 9)
                //{
                //    filenameM = string.Format("{0}_{1}_{2}_{3}_{4}", arg[1], arg[7], arg[6], arg[5], arg[4]);
                //    filepathM = arg[8];
                //}
                //else
                //{ return; }
                foreach (var a in arg[3].Split(',').ToList())
                {
                    BlockList.Add(a);
                }
                //List<PluginVO> Plugins = Connector.GetPluginControlList();
                //foreach (PluginVO plugin in Plugins)
                //{
                //    // Plugin.xml Menu Title Attribute 정의한 텍스트 
                //    if (plugin.MenuTitle == "11.간섭체크(정적/동적/러그_와이어)")
                //    {
                //        ((DevExpress.XtraBars.BarCheckItem)plugin.MenuButtonControl).PerformClick();
                //    }
                //}

                Connector.OnFinishedClashTestEvent += Connector_AUTO_FinishedClashTestEvent;
                timer_AUTO.Enabled = false;
                timer_AUTO.Tick += timer_AUTO_Tick;

                lugtime.Restart();

                Console.WriteLine("Start, button2_Click");

                ListViewItem bb = new ListViewItem();
                List<string> blocks = new List<string>();
                foreach (var a in arg[3].Split(',').ToList())
                {
                    bb.SubItems[0].Text = a;
                    blocks.Add(a);
                }
                watListView1.Items.Add(bb);
                //Connector.AddDocuments(blocks.ToArray());

                foreach (var a in blocks)
                {
                    ClashData data = new ClashData();
                    data.BlockName = a;
                    Clash.Add(data);
                }

                AUTOClashRun();
                Console.WriteLine("End, button2_Click");
            }
            catch(Exception ea) { MessageBox.Show(ea.Message); }

        }
        private void getSystemInfo()
        {
            ProgBar pBar = new ProgBar(1, String.Format("Memory Checking... (Total {0})", 1)); //
            pBar.Show();
            pBar.IncreaseValue_ChangeTitle(String.Format("Memory Checking... (Total 1/{0})", 1));

            string outputFile = Path.Combine(_tempDir2, "SYSTEMINFO.TXT");

            string runCMDStr = String.Format(@"SYSTEMINFO > {0}", outputFile);
            Console.WriteLine("***runCMDStr:{0}", runCMDStr);
            RunCMDOnly(runCMDStr, true, true);

            if (File.Exists(outputFile) == true)
            {
                string writeTxt = "";
                string _memoryTxt = "";
                string[] readLines = File.ReadAllLines(outputFile, Encoding.Default);
                foreach (string readLine in readLines)
                {
                    writeTxt = String.Format("{0}\r\n{1}", writeTxt, readLine);
                    if (readLine.IndexOf(@"총 실제 메모리") >= 0)
                    {
                        _memoryTxt = String.Format("{0}", readLine.Substring(readLine.IndexOf(":") + 1).Trim());
                        _memoryTxt = _memoryTxt.Replace(",", "");
                        _memoryTxt = _memoryTxt.Replace("MB", "");
                        try { _dMemory = double.Parse(_memoryTxt); }
                        catch { }
                    }
                }
                //label12.Text = String.Format(@"총 실제 메모리 : {0} MB", _dMemory); //writeTxt
            }

            pBar.Close();
        }

        //CMD 실행
        void RunCMDOnly(string runFile, bool consoleYes, bool waitForExit)
        {
            //
            try
            {
                //string execStmt = String.Format(@" /C ""{0}""", runFile);
                string execStmt = String.Format(@" /C " + @"{0}", runFile);
                System.Diagnostics.ProcessStartInfo procStartInfo = new System.Diagnostics.ProcessStartInfo("cmd.exe", execStmt);

                procStartInfo.RedirectStandardOutput = false; //true; (중요)
                procStartInfo.UseShellExecute = false;

                procStartInfo.CreateNoWindow = consoleYes; //true->WINDOWS(GUI) BASE, false->CONSOLE BASE

                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo = procStartInfo;
                proc.Start();
                if (waitForExit == true) { proc.WaitForExit(); }
                proc.Close();
            }
            catch
            {
            }
            //
            return;
        }

        private void InitDrive()
        {
            Console.WriteLine("Start, InitDrive");

            DirectoryInfo dir = new DirectoryInfo(_originPath);
            DirectoryInfo[] sDir = dir.GetDirectories();
            Console.WriteLine("@0");

            List<string> itemlist = new List<string>();
            itemlist.Add(""); //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            foreach (DirectoryInfo dd in sDir)
            {
                if (dd.ToString() == _shipNo) { continue; }
                //cbBox1.Items.Add(dd.Name);
                string strTmp = Regex.Replace(dd.Name, @"\D", "");
                if (dd.Name.Length == 6 && strTmp.Length != 0)
                {
                    itemlist.Add(dd.Name);
                }
            }
            itemlist.Sort();
            cbBox1.DataSource = itemlist;
            try
            {
                //cbBox1.SelectedItem = cbBox1.Items[cbBox1.Items.Count - 1]; 
                cbBox1.SelectedItem = cbBox1.Items[0];
            }
            catch { }

            Console.WriteLine("@1");

            cbBox1.SelectedIndexChanged += new EventHandler(cbBox1_SelectedIndexChanged);

            Console.WriteLine("End, InitDrive");
        }

        public bool LicenseCheck()
        {
            // License Check - Start

            var checkTime = DateTime.Now;

            bool bOK = true;

            decimal timeTemp = decimal.Parse(String.Format("{0:D4}{1:D2}{2:D2}{3:D2}{4:D2}{5:D2}.{6}", checkTime.Year, checkTime.Month, checkTime.Day, checkTime.Hour, checkTime.Minute, checkTime.Second, checkTime.Millisecond));

            Console.WriteLine("_allowTimeMAX:{0}, _allowTimeMIN:{1}", _allowTimeMAX, _allowTimeMIN);
            if (timeTemp < _allowTimeMIN)
            {
                bOK = false;
            }
            if (timeTemp > _allowTimeMAX)
            {
                bOK = false;
            }
            // License Check - End
            return bOK;
        }
        private void btnLoadModel_Click(object sender, EventArgs e)
        {
            GC.Collect(); GC.WaitForPendingFinalizers(); //Garbage

            //
            string vizXMLFile = Path.Combine(_tempDir2, String.Format(@"vizXMLFile.VIZXML"));
            Console.WriteLine("vizXMLFile:{0}", vizXMLFile);

            _selectedItems.Clear();

            ArrayList writeLines = new ArrayList();

            for (int i = 0; i < watListView1.SelectedItems.Count; i++)
            {
                _selectedItems.Add(watListView1.SelectedItems[i].SubItems[0].Text);
            }

            try
            {
                FileStream fileStreamOutput = new FileStream(vizXMLFile, FileMode.Create);
                fileStreamOutput.Seek(0, SeekOrigin.Begin);

                byte[] info;

                writeLines.Add(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
                writeLines.Add(@"<VIZXML>");
                writeLines.Add(String.Format(@"<Model Name=""{0}"" SkipBrokenLinks=""True"">", _shipNo));
                Connector.ShowWaitDialogWithText(true,"Hole-Split","변환중");              
                if(allblock)
                {
                    for (int i = 0; i < watListView1.SelectedItems.Count; i++)
                    {
                        string b = watListView1.SelectedItems[i].SubItems[0].Text;
                        string vizFile = Path.Combine(_selectedSHIPPath, String.Format(@"{0}\{1}_정적.viz", b, b));//re
                        FileInfo file = new FileInfo(vizFile);

                        string block = "";
                        bool isfirst = true;
                        foreach (var a in tree)
                        {
                            //if (a.Contains("/" + b + "/"))
                            //{
                            if (a.Split('/').Last() == b && isfirst)
                            {
                                if (isfirst)
                                {
                                    block = a.Split('/').ToList()[2];
                                    Console.WriteLine(block);
                                    isfirst = false;
                                }
                            }
                        }
                        string path = _originPath; string filepath = "";
                        bool first = true;
                        foreach (var a in tree)
                        {
                            if (a.Split('/').Last() == b && first)
                            {
                                if (a.Split('/')[a.Split('/').Count() - 2].Length == 4 || a.Split('/')[a.Split('/').Count() - 2] == _shipNo)
                                {
                                    filepath = a;
                                    first = false;
                                }
                            }
                        }
                        filepath = filepath.Substring(_shipNo.Length + 1);
                        filepath = filepath.Replace('/', '\\');
                        if (file.Exists)
                        {
                            writeLines.Add(string.Format("<Node Name=\"{3}\" ExtLinkNode=\"{4}\\{0}\\{1}\\{1}_정적.viz:{2}\" HideAndLock=\"False\" UncheckToUnload=\"False\" Type=\"Assembly\"/>",
                              _shipNo, block, filepath, b, path));
                        }
                        else
                        {
                            writeLines.Add(string.Format("<Node Name=\"{3}\" ExtLinkNode=\"{4}\\{0}\\{1}\\{1}.rev:{2}\" HideAndLock=\"False\" UncheckToUnload=\"False\" Type=\"Assembly\"/>",
  _shipNo, block, filepath, b, path));
                        }
                    }
                }
                else
                {
                    //2021-01-14 Select -> 전체 블록으로 변경
                    for (int i = 0; i < watListView1.Items.Count; i++)
                    {
                        string selectedItemLAST = watListView1.Items[i].SubItems[0].Text;

                        string vizFile = Path.Combine(_selectedSHIPPath, String.Format(@"{0}\{1}_정적.viz", selectedItemLAST, selectedItemLAST));//re
                        FileInfo file = new FileInfo(vizFile);
                        if (!file.Exists)
                        {
                            vizFile = Path.Combine(_selectedSHIPPath, String.Format(@"{0}\{1}.rev", selectedItemLAST, selectedItemLAST));//re
                        }
                        Connector.UpdateWaitDialogDescription(string.Format("{0}", selectedItemLAST));

                        ////FileInfo FI = new FileInfo(target);
                        ////if (!FI.Exists)
                        ////{
                        ////CurveSplit curveSplit = new CurveSplit();
                        ////curveSplit.Split(vizFile, target, 3, 2, 0.1); //블록 파일 변경가능성 --> 계속 재 생성 및 덮어씀;;;

                        //string target = Path.Combine(_selectedSHIPPath, String.Format(@"{0}\{1}_HoleSplit.rev", selectedItemLAST, selectedItemLAST));
                        //runcmd(vizFile, target);//rev SKIP


                        writeLines.Add(String.Format(@"<Node Name=""{0}"" ExtLinkFile=""{1}"" HideAndLock=""False"" Type=""Assembly""/>", selectedItemLAST, vizFile));
                        string time2 = string.Format("{0}시{1}분{2}초", DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00")
    , DateTime.Now.Second.ToString("00"));
                        //File.AppendAllText(logfilePath, string.Format("[{0} {1}] {2}블록 HoleSplit", DateTime.Now.ToShortDateString(), time2, a) + Environment.NewLine);
                        //}
                        //else
                        //{
                        //    writeLines.Add(String.Format(@"<Node Name=""{0}"" ExtLinkFile=""{1}"" HideAndLock=""False"" Type=""Assembly""/>", selectedItemLAST, vizFile));
                        //}
                    }
                }
                Connector.ShowWaitDialog(false);
                if (s_attFileNames.StartsWith("@") == true) { s_attFileNames = s_attFileNames.Substring(1); }

                writeLines.Add(@"</Model>");
                writeLines.Add(@"</VIZXML>");

                foreach (string writeLine in writeLines)
                {
                    //info = System.Text.Encoding.Default.GetBytes(writeLine + "\r\n"); //ANSI
                    info = new UTF8Encoding(true).GetBytes(writeLine + "\r\n"); //UTF-8
                    fileStreamOutput.Write(info, 0, info.Length);
                }

                fileStreamOutput.Flush();
                fileStreamOutput.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            //

            CLASH_MODEL_FILE = vizXMLFile;
            string time = string.Format("{0}시{1}분{2}초", DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00")
, DateTime.Now.Second.ToString("00"));
            File.AppendAllText(logfilePath, string.Format("[{0} {1}] 블록 Load 시작", DateTime.Now.ToShortDateString(), time) + Environment.NewLine);

            Connector.OpenDocument(vizXMLFile);
            string lugcheck = Path.Combine(_originPath, _shipNo, "LUG_WIRE", "LUG_WIRE.REV");
            if (!new FileInfo(lugcheck).Exists)
            {
                lugwire = false;
            }
            if (lugwire)
            {
                //Lug Wire 추가
                string lugPath = Path.Combine(_originPath, _shipNo, "LUG_WIRE", "LUG_WIRE.viz");
                if (!new FileInfo(lugPath).Exists)
                {
                    lugPath = Path.Combine(_originPath, _shipNo, "LUG_WIRE", "LUG_WIRE.REV");
                }
                List<string> lugadd = new List<string>() { lugPath };
                Connector.AddDocuments(lugadd.ToArray());
            }

            Connector.SetUncheckToUnload(false);
            time = string.Format("{0}시{1}분{2}초", DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00")
, DateTime.Now.Second.ToString("00"));
            File.AppendAllText(logfilePath, string.Format("[{0} {1}] 블록 Load 끝", DateTime.Now.ToShortDateString(), time) + Environment.NewLine);
        }
        public void runcmd(string rev, string hole)
        {
            ProcessStartInfo proinfo = new ProcessStartInfo();
            Process pro = new Process();

            //실행할 파일명 입력 --cmd
            proinfo.FileName = @"cmd";
            //cmd창 띄우기 -- true(띄우지 않기.)false(띄우기)
            proinfo.CreateNoWindow = true;
            proinfo.UseShellExecute = false;
            //cmd 데이터 받기
            proinfo.RedirectStandardOutput = true;
            //cmd 데이터 보내기
            proinfo.RedirectStandardInput = true;
            //cmd 오류내용 받기
            proinfo.RedirectStandardError = false;

            pro.StartInfo = proinfo;
            pro.Start();
            string company = "IFG"; string curvesplit = @"Q:\env\Vitesse\Infoget\autoAP\IGRevGenerator.exe";
            if (Directory.Exists(@"M:\HMD") == true)
            {
                company = "HMD";
            }
            //cmd에 보낼 명령어를 입력 합니다.
            if(company =="IFG")
            {
                curvesplit = @"D:\IGRevGenerator\IGRevGenerator\IGRevGeneratorTest\bin\Debug\IGRevGenerator.exe";
            }
            Console.WriteLine(curvesplit);
            pro.StandardInput.Write(string.Format("{0} {1} {2} {3} {4} {5}", curvesplit,rev,hole,"3","2","0.1")
                + Environment.NewLine);
            pro.StandardInput.Close();

            ////결과 값을 리턴 받습니다.
            string resultValue = pro.StandardOutput.ReadToEnd();
            pro.WaitForExit();
            pro.Close();

            //결과 값을 확인 합니다.
            Console.WriteLine(resultValue);

            //return resultValue;
        }
        int blockindex = 0;
        private void btnShowModel_Click(object sender, EventArgs e)
        {
            DateTime ss = DateTime.Now;
            Connector.ShowObject(0, true);
            Connector.ViewInitPosition();
            DateTime ed = DateTime.Now;
            TimeSpan ad = ed - ss;
            TotalLoad.Add(ad);//w
            //Lug 는 비활성화
            int nodeindex = 0;
            foreach(var a in Connector.GetAllObjects())
            {
                if (a.NodeName.Contains("LUG_WIRE.REV"))
                {
                    nodeindex = a.Index;
                }
                if (a.NodeName.Equals(_shipNo))
                {
                    blockindex = a.Index;
                }
            }
            if(nodeindex!=0)
            {
                Connector.ShowObject(nodeindex,false);
            }

            ObjectPropertyVO prop = Connector.GetObjectProperty(0, false);

            SHIP_MIN_Z = Convert.ToSingle(prop.MinPoint.Z);
            SHIP_MAX_Z = Convert.ToSingle(prop.MaxPoint.Z);

        }

        private void btnClashTest_Click(object sender, EventArgs e)
        {
            //Connector.ShowObject(0, false);

            Connector.ShowWaitDialogWithText(true, "간섭검사", "검사를 위한 정보 수집중...");

            string time = string.Format("{0}시{1}분{2}초", DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00")
, DateTime.Now.Second.ToString("00"));
            File.AppendAllText(logfilePath, string.Format("[{0} {1}] 간섭검사 정보 수집", DateTime.Now.ToShortDateString(), time) + Environment.NewLine);

            TaskList = new List<ClashTaskMultiVO>();
            lvList.Items.Clear();

            Connector.SetObjectsColor(new int[] { blockindex }, Color.White);

            int nStepDistance = Convert.ToInt32(txtStepDistance.Text);


            List<NodeVO> items = Connector.GetChildObjects(blockindex, ChildrenTypes.Children);
            //List<string> unloadnode = new List<string>();
            //for (int i = 0; i < items.Count; i++) //아이템 있는지 확인
            //{
            //    if(items[i].IsUnloadedNode==true)
            //    {
            //        unloadnode.Add(items[i].NodeName);
            //    }
            //}
            //if(unloadnode.Count!=0)
            //{
            //    MessageBox.Show(unloadnode.Count.ToString());
            //    return;
            //}
            //StreamWriter sw = new StreamWriter(@"C:\temp\moos.txt");//lll

            Console.WriteLine("items.Count:{0}", items.Count);
            for (int i = 0; i < items.Count; i++) //
            {
                Console.WriteLine("items[{0}].NodeName.ToUpper():{1}", i, items[i].NodeName.ToUpper());

                //if (i == 0) { continue; } //very very important

                Connector.UpdateWaitDialogDescription(string.Format("분석중 : [{0}/{1}] : ({2})", i + 1, items.Count, items[i].NodeName));
                time = string.Format("{0}시{1}분{2}초", DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00")
, DateTime.Now.Second.ToString("00"));
                File.AppendAllText(logfilePath, string.Format("[{0} {1}]분석중 : [{2}/{3}] : ({4})", DateTime.Now.ToShortDateString(), time, i + 1, items.Count, items[i].NodeName) + Environment.NewLine);

                //Console.WriteLine(" [{0}]", items[i].IsUnloadedNode);

                if (items[i].IsUnloadedNode == true) { continue; }

                if (_selectedItems.Contains(items[i].NodeName) == false) { continue; } //to be deleted

                Connector.EnableRender(false);
                try
                {
                    Connector.SetObjectsColor(new int[] { items[i].Index }, Color.Orange);
                Connector.SetObjectsColor(new int[] { items[i - 1].Index }, Color.White); }
                catch { }
                Connector.EnableRender(true);

                Console.WriteLine(" ~@1");
                // 고정블록 검색
                List<float> box = new List<float>();
                ObjectPropertyVO mProp = Connector.GetObjectProperty(items[i].Index, false); //
                box.Add(Convert.ToSingle(mProp.MinPoint.X) - 1000.0f);
                box.Add(Convert.ToSingle(mProp.MinPoint.Y) - 1000.0f);
                box.Add(Convert.ToSingle(mProp.MinPoint.Z) - 2000.0f);
                //box.Add(SHIP_MIN_Z);
                box.Add(Convert.ToSingle(mProp.MaxPoint.X) + 1000.0f);
                box.Add(Convert.ToSingle(mProp.MaxPoint.Y) + 1000.0f);
                box.Add(Convert.ToSingle(mProp.MaxPoint.Z) + 1000.0f);
                //box.Add(SHIP_MAX_Z);
                Console.WriteLine(" ~mProp.MinPoint.X:{0}, mProp.MinPoint.Y:{1}, mProp.MinPoint.Z:{2}", mProp.MinPoint.X, mProp.MinPoint.Y, mProp.MinPoint.Z);
                Console.WriteLine(" ~mProp.MaxPoint.X:{0}, mProp.MaxPoint.Y:{1}, mProp.MaxPoint.Z:{2}", mProp.MaxPoint.X, mProp.MaxPoint.Y, mProp.MaxPoint.Z);
                //Console.WriteLine(" ~~~SHIP_MAX_Z:{0}", SHIP_MAX_Z);

                List<int> searchItems = Connector.GetObjectsInArea(box.ToArray(), new int[] { }, CrossBoundBox.IncludingPart);//IncludingPart

                //multiVO.FIXED_BLOCK = GetFixedBlockNames(items, i);

                Console.WriteLine(" ~@2");

                string BLOCK_B = items[i].NodeName;
                Dictionary<string, string> SourceMap = GetFixedBlockNamesMap(items, i);
                Dictionary<string, string> ResultMap = new Dictionary<string, string>();

                Dictionary<string, string> fixedBlockDic = new Dictionary<string, string>();
                fixedBlockDic[BLOCK_B] = "";

                for (int j = 0; j < searchItems.Count; j++)
                {
                    string path = Connector.GetNodePath(searchItems[j]);
                    string[] NODES = path.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);

                    if (NODES[blockindex] == BLOCK_B) { continue; }

                    if (SourceMap.ContainsKey(NODES[blockindex]) == true)
                    {
                        if (ResultMap.ContainsKey(NODES[blockindex]) == false)
                        {
                            ResultMap.Add(NODES[blockindex], NODES[blockindex]);
                        }
                    }
                }
                ////인접블록 강제 부여
                //ResultMap.Clear();
                //StreamReader sr = new StreamReader(@"C:\Temp\mooon.txt");
                //List<string> readdata = new List<string>();
                //string line;
                //while ((line = sr.ReadLine()) != null)
                //{
                //    readdata.Add(line);
                //}
                //sr.Close();
                //foreach (var a in readdata)
                //{
                //    if (a == "")
                //        continue;


                //    if (items[i].NodeName.Contains(a.Split('@').First()))
                //    {
                //        List<string> fix = a.Split('@').Last().Split('#').ToList();
                //        foreach (var b in fix)
                //        {
                //            if (b != "")
                //            {
                //                ResultMap.Add(b, b);
                //            }
                //        }
                //    }
                //}

                //

                string FIXED_BLOCK_STR = String.Empty;
                foreach (string item in ResultMap.Keys)
                {
                    if (String.IsNullOrEmpty(FIXED_BLOCK_STR) == false)
                        FIXED_BLOCK_STR += ", ";

                    FIXED_BLOCK_STR += item;
                }

                ////인접블록 save
                //string wr = items[i].NodeName + "@@@" + FIXED_BLOCK_STR;
                //sw.WriteLine(wr);


                List<float> LBD = GetBlockHeight(items[i].Index);
                Console.WriteLine(" fHeight:{0}", LBD[2]);
                File.AppendAllText(logfilePath, string.Format("{0}  높이 : {1}",BLOCK_B,LBD[2]+1000));
                // 간섭검사 VO 생성
                ClashTaskVO taskGroup = new ClashTaskVO();
                ClashTaskVO taskModel = new ClashTaskVO();
                ClashTaskVO taskMove = new ClashTaskVO();

                // 간섭검사 기준 데이터 설정
                taskGroup.CtType = CtType.CtTypeG2G; //정적간섭검사(그룹)
                taskModel.CtType = CtType.CtTypeSelf; //정적간섭검사(모델) ->추가 2018.12.10.
                taskMove.CtType = CtType.CtTypeMoving; //동적간섭검사

                // 그룹 설정 : False (간섭검사 UI와 분리) / True (간섭검사 UI와 연동)
                taskGroup.bInitTranslationValue = true;
                taskModel.bInitTranslationValue = true;
                taskMove.bInitTranslationValue = true;

                // 보이는 모델만
                taskGroup.VisibleOnly = true;
                taskModel.VisibleOnly = true;
                taskMove.VisibleOnly = true;

                // 공간으로 고정블록/탑재블록 할당을 위해서 추후 설정하며,
                // 로드/언로드로 인해서 검사 시작전 설정이 되어야 함
                //task.FixedGroupNodes.AddRange(GetFixedBlock(items, i));
                //task.MovingGroupNodes.Add(items[i]);

                // 이동검사 정보 추가
                if (taskMove.CtType == CtType.CtTypeMoving)
                {
                    string direction = watListView1.Items[i].SubItems[7].Text;
                    if (direction == "Z"|| direction == "Z+")
                    {
                        // 사용자가 정의한 값을 사용하는 경우 True / Min & Max를 활용한 Route 생성은 False
                        taskMove.UseCustomMovingValue = true;

                        int nHeight = Convert.ToInt32(LBD[2]) + 1000; // 임의로 1m 가량 추가
                        int nStep = nHeight / nStepDistance; //
                        Console.WriteLine(" nStep:{0}, nHeight:{1}, nStepDistance:{2}", nStep, nHeight, nStepDistance);

                        for (int j = nStep; j > 0; j--)
                        {
                            taskMove.MovingRouteList.Add(new MovingRouteVO(0, 0, (j * nStepDistance)));
                        }
                    }
                    else if (direction == "X-")
                    {
                        // 사용자가 정의한 값을 사용하는 경우 True / Min & Max를 활용한 Route 생성은 False
                        taskMove.UseCustomMovingValue = true;

                        //int nHeight = Convert.ToInt32(LBD[0]) + 1000; // 임의로 1m 가량 추가
                        int nHeight = int.Parse(watListView1.Items[i].SubItems[8].Text.Replace("mm", ""));

                        int nStep = nHeight / nStepDistance; //
                        Console.WriteLine(" nStep:{0}, nHeight:{1}, nStepDistance:{2}", nStep, nHeight, nStepDistance);

                        for (int j = nStep; j > 0; j--)
                        {
                            taskMove.MovingRouteList.Add(new MovingRouteVO((j * nStepDistance), 0, 0));
                        }

                        int lastSTEP = nStepDistance * nStep;
                        nHeight = Convert.ToInt32(LBD[2]) + 1000; // 임의로 1m 가량 추가
                        nStep = nHeight / nStepDistance; //
                        Console.WriteLine(" nStep:{0}, nHeight:{1}, nStepDistance:{2}", nStep, nHeight, nStepDistance);

                        for (int j = nStep; j > 0; j--)
                        {
                            taskMove.MovingRouteList.Add(new MovingRouteVO(lastSTEP, 0, (j * nStepDistance)));
                        }
                    }
                    else if (direction == "X+")
                    {
                        // 사용자가 정의한 값을 사용하는 경우 True / Min & Max를 활용한 Route 생성은 False
                        taskMove.UseCustomMovingValue = true;

                        //int nHeight = Convert.ToInt32(LBD[0]) + 1000; // 임의로 1m 가량 추가
                        int nHeight = int.Parse(watListView1.Items[i].SubItems[8].Text.Replace("mm", ""));
                        int nStep = nHeight / nStepDistance; //
                        Console.WriteLine(" nStep:{0}, nHeight:{1}, nStepDistance:{2}", nStep, nHeight, nStepDistance);

                        for (int j = nStep; j > 0; j--)
                        {
                            taskMove.MovingRouteList.Add(new MovingRouteVO(-(j * nStepDistance), 0, 0));
                        }

                        int lastSTEP = nStepDistance * nStep;
                        nHeight = Convert.ToInt32(LBD[2]) + 1000; // 임의로 1m 가량 추가
                        nStep = nHeight / nStepDistance; //
                        Console.WriteLine(" nStep:{0}, nHeight:{1}, nStepDistance:{2}", nStep, nHeight, nStepDistance);

                        for (int j = nStep; j > 0; j--)
                        {
                            taskMove.MovingRouteList.Add(new MovingRouteVO(-lastSTEP, 0, (j * nStepDistance)));
                        }
                    }
                    else if (direction == "Y-")
                    {
                        // 사용자가 정의한 값을 사용하는 경우 True / Min & Max를 활용한 Route 생성은 False
                        taskMove.UseCustomMovingValue = true;

                        //int nHeight = Convert.ToInt32(LBD[1]) + 1000; // 임의로 1m 가량 추가
                        int nHeight = int.Parse(watListView1.Items[i].SubItems[8].Text.Replace("mm", ""));
                        int nStep = nHeight / nStepDistance; //
                        Console.WriteLine(" nStep:{0}, nHeight:{1}, nStepDistance:{2}", nStep, nHeight, nStepDistance);

                        for (int j = nStep; j > 0; j--)
                        {
                            taskMove.MovingRouteList.Add(new MovingRouteVO(0, (j * nStepDistance), 0));
                        }

                        int lastSTEP = nStepDistance * nStep;
                        nHeight = Convert.ToInt32(LBD[2]) + 1000; // 임의로 1m 가량 추가
                        nStep = nHeight / nStepDistance; //
                        Console.WriteLine(" nStep:{0}, nHeight:{1}, nStepDistance:{2}", nStep, nHeight, nStepDistance);

                        for (int j = nStep; j > 0; j--)
                        {
                            taskMove.MovingRouteList.Add(new MovingRouteVO(0, lastSTEP, (j * nStepDistance)));
                        }
                    }
                    else if (direction == "Y+")
                    {
                        // 사용자가 정의한 값을 사용하는 경우 True / Min & Max를 활용한 Route 생성은 False
                        taskMove.UseCustomMovingValue = true;

                        //int nHeight = Convert.ToInt32(LBD[1]) + 1000; // 임의로 1m 가량 추가
                        int nHeight = int.Parse(watListView1.Items[i].SubItems[8].Text.Replace("mm", ""));
                        int nStep = nHeight / nStepDistance; //
                        Console.WriteLine(" nStep:{0}, nHeight:{1}, nStepDistance:{2}", nStep, nHeight, nStepDistance);

                        for (int j = nStep; j > 0; j--)
                        {
                            taskMove.MovingRouteList.Add(new MovingRouteVO(0, -(j * nStepDistance), 0));
                        }

                        int lastSTEP = nStepDistance * nStep;
                        nHeight = Convert.ToInt32(LBD[2]) + 1000; // 임의로 1m 가량 추가
                        nStep = nHeight / nStepDistance; //
                        Console.WriteLine(" nStep:{0}, nHeight:{1}, nStepDistance:{2}", nStep, nHeight, nStepDistance);

                        for (int j = nStep; j > 0; j--)
                        {
                            taskMove.MovingRouteList.Add(new MovingRouteVO(0, -lastSTEP, (j * nStepDistance)));
                        }
                    }
                    else
                    {
                        MessageBox.Show(string.Format("탑재 방향 ERROR{0}", direction));
                        return;
                    }
                    if(lastSTEP)
                    {
                        taskMove.MovingRouteList.Add(new MovingRouteVO(0, 0, 0));
                    }
                }

                ListViewItem lviMove = new ListViewItem(new string[] {
                    string.Format("{0:D2}", i)
                    , items[i].NodeName
                    , string.Format("{0}", Math.Round(LBD[2], 2))
                    , FIXED_BLOCK_STR
                    , string.Format("{0}", taskMove.MovingRouteList.Count)
                    , "N/A"
                    , "N/A"
                });
                lvList.Items.Add(lviMove);

                // 간섭검사 환경 설정
                // 어셈블리 단위 설정
                taskGroup.CheckFixedGroupAsm = true;
                taskGroup.CheckMovingGroupAsm = true;

                taskModel.CheckFixedGroupAsm = true;
                taskModel.CheckMovingGroupAsm = true;

                taskMove.CheckFixedGroupAsm = true;
                taskMove.CheckMovingGroupAsm = true;

                // 근접허용범위
                taskGroup.bClashToleranceRange = true;
                taskGroup.bRange = true;
                taskGroup.RangeTolerance = Convert.ToInt32(textBoxNear.Text); //2.0f;

                taskModel.bClashToleranceRange = true;
                taskModel.bRange = true;
                taskModel.RangeTolerance = Convert.ToInt32(textBoxNear.Text); //2.0f;


                
                taskMove.bClashToleranceRange = true;
                taskMove.bRange = true;
                taskMove.RangeTolerance = Convert.ToInt32(textBoxNear.Text); //2.0f;
                

               

                // 접촉허용오차
                taskGroup.bClashCalibrationTolerance = true;
                taskGroup.bPenet = true;
                taskGroup.ContractTolerance = Convert.ToInt32(textBoxContact.Text); //1.0f;

                taskModel.bClashCalibrationTolerance = true;
                taskModel.bPenet = true;
                taskModel.ContractTolerance = Convert.ToInt32(textBoxContact.Text); //1.0f;               

                taskMove.bClashCalibrationTolerance = true;
                taskMove.bPenet = true;
                taskMove.ContractTolerance = Convert.ToInt32(textBoxContact.Text); //1.0f;
                

                // 간섭제외 끝레벨
                taskGroup.BottomLevel = int.Parse(textBoxExceptLevel.Text); //2;
                taskModel.BottomLevel = int.Parse(textBoxExceptLevel.Text); //2;
                taskMove.BottomLevel = int.Parse(textBoxExceptLevel.Text); //2;

                Console.WriteLine(" ~@3");

                //간섭검사 시작
                //Connector.Clash_StartCheck(task);
                bool bIsSelected = false;
                foreach(ListViewItem li in watListView1.SelectedItems)
                {
                    if(items[i].NodeName==li.SubItems[0].Text)
                    {
                        bIsSelected=true;
                    }
                }

                if(!bIsSelected)
                {
                    continue;
                }
                // 별도 쓰레드 처리를 위한 작업 목록에 작업 생성 및 추가
                ClashTaskMultiVO multiVOModel = new ClashTaskMultiVO();
                multiVOModel.Connector = Connector;
                multiVOModel.TaskVO = taskModel;
                multiVOModel.ListItem = null; // lviGroup;
                multiVOModel.MOVING_BLOCK = items[i].NodeName;
                multiVOModel.FIXED_BLOCK = fixedBlockDic.Keys.ToList();
                TaskList.Add(multiVOModel);

                if (i != 0)
                {
                    ClashTaskMultiVO multiVOGroup = new ClashTaskMultiVO();
                    multiVOGroup.Connector = Connector;
                    multiVOGroup.TaskVO = taskGroup;
                    multiVOGroup.ListItem = null; // lviGroup;
                    multiVOGroup.MOVING_BLOCK = items[i].NodeName;
                    //multiVOGroup.LUGWIRE = "LUG_WIRE_" + items[i].NodeName;//그룹제외
                    multiVOGroup.FIXED_BLOCK = ResultMap.Keys.ToList();
                    TaskList.Add(multiVOGroup);

                    ClashTaskMultiVO multiVOMove = new ClashTaskMultiVO();
                    multiVOMove.Connector = Connector;
                    multiVOMove.TaskVO = taskMove;
                    multiVOMove.ListItem = lviMove;
                    multiVOMove.MOVING_BLOCK = items[i].NodeName;
                    multiVOMove.FIXED_BLOCK = ResultMap.Keys.ToList();
                    if (!Stand.Checked)
                    {
                        if (!exceptBlock.Contains(items[i].NodeName))
                        {
                            TaskList.Add(multiVOMove);
                        }
                    }
                }

                Console.WriteLine(" ~@4");
            }

            // 모델 언로드
            List<int> IDS = new List<int>();
            for (int i = 0; i < items.Count; i++)
            {
                IDS.Add(items[i].Id);
            }
            Connector.UnloadMultiNode(IDS.ToArray());

            Connector.ShowWaitDialog(false);
        }
        private void btnStartClashTest_Click(object sender, EventArgs e)
        {
            if (TaskList == null) return;
            if (TaskList.Count == 0) return;

            Connector.OnFinishedClashTestEvent += Connector_OnFinishedClashTestEvent;
            // 간섭검사 쓰레드 시작
            timerClash.Enabled = true;
        }

        private List<float> GetBlockHeight(int NodeIndex)
        {
            ObjectPropertyVO prop = Connector.GetObjectProperty(NodeIndex, false);

            float MaxZ = Convert.ToSingle(prop.MaxPoint.Z);
            float MinZ = Convert.ToSingle(prop.MinPoint.Z);

            float D = MaxZ - MinZ;

            float MaxX = Convert.ToSingle(prop.MaxPoint.X);
            float MinX = Convert.ToSingle(prop.MinPoint.X);

            float L = MaxX - MinX;

            float MaxY = Convert.ToSingle(prop.MaxPoint.Y);
            float MinY = Convert.ToSingle(prop.MinPoint.Y);

            float B = MaxY - MinY;

            List<float> LBD = new List<float>() { L, B, D };

            return LBD;
        }

        private List<NodeVO> GetFixedBlock(List<NodeVO> items, int index)
        {
            List<NodeVO> blocks = new List<NodeVO>();

            for (int i = 0; i < index; i++)
            {
                blocks.Add(items[i]);
            }

            return blocks;
        }

        private List<string> GetFixedBlockNames(List<NodeVO> items, int index)
        {
            List<string> blocks = new List<string>();

            for (int i = 0; i < index; i++)
            {
                blocks.Add(items[i].NodeName);
            }

            return blocks;
        }

        private Dictionary<string, string> GetFixedBlockNamesMap(List<NodeVO> items, int index)
        {
            Dictionary<string, string> blocks = new Dictionary<string, string>();

            for (int i = 0; i < index; i++)
            {
                blocks.Add(items[i].NodeName, items[i].NodeName);
            }

            return blocks;
        }

        private string GetPrevBlocks(List<NodeVO> items, int index)
        {
            string ret = String.Empty;

            for (int i = 0; i < index; i++)
            {
                if (String.IsNullOrEmpty(ret) == false)
                    ret += ", ";

                ret += items[i].NodeName;
            }

            return ret;
        }

        private void timerClash_Tick(object sender, EventArgs e)
        {
            Console.WriteLine("Start, timerClash_Tick");
            File.AppendAllText(logfilePath, "Start, timerClash_Tick" + Environment.NewLine);

            timerClash.Enabled = false;

            ClashTaskVO taskVO = null;

            Console.WriteLine("TaskList.Count:{0}", TaskList.Count);
            for (int i = 0; i < TaskList.Count; i++)
            {
                Console.WriteLine("#i:{0}", i);

                ClashTaskMultiVO multiVO = TaskList[i];

                if (multiVO.IsCompleted == true) { continue; } // 완료된 항목은 제외

                // Root 블록
                List<NodeVO> CHILD = Connector.GetChildObjects(blockindex, ChildrenTypes.Children);
                //List<int> NODE_IDS = new List<int>();
                Dictionary<string, int> LOADED_BLOCKS = new Dictionary<string, int>();
                for (int j = 0; j < CHILD.Count; j++)
                {
                    if (CHILD[j].IsUnloadedNode == true) { continue; }

                    //NODE_IDS.Add(CHILD[j].Id);

                    if (LOADED_BLOCKS.ContainsKey(CHILD[j].NodeName) == false)
                        LOADED_BLOCKS.Add(CHILD[j].NodeName, CHILD[j].Id);
                }
                File.AppendAllText(logfilePath, "Start, TASK" + Environment.NewLine);
                // 언로드
                //Connector.UnloadMultiNode(NODE_IDS.ToArray());
                // 검사 대상 Task
                if (multiVO.IsTesting == false)
                {
                    multiVO.IsTesting = true;   // 진행중으로 상태 변경

                    DateTime ss = DateTime.Now;
                    // 이동블록 로드
                    List<NodeVO> SearchMovingBlocks = GetNodeSearch(multiVO.MOVING_BLOCK);
                    NodeVO testnode = Connector.GetObject(SearchMovingBlocks[0].Index);
                    File.AppendAllText(logfilePath, "Moving Node Index : " + SearchMovingBlocks[0].Index + Environment.NewLine);
                    File.AppendAllText(logfilePath, "Moving Node Name : " + testnode.NodeName + Environment.NewLine);
                    Connector.ShowObject(SearchMovingBlocks[0].Index, true);
                    if (LOADED_BLOCKS.ContainsKey(SearchMovingBlocks[0].NodeName) == true)
                        LOADED_BLOCKS.Remove(SearchMovingBlocks[0].NodeName);

                    // 고정블록 로드
                    foreach (string item in multiVO.FIXED_BLOCK)
                    {
                        List<NodeVO> SearchFixedBlocks = GetNodeSearch(item);
                        NodeVO testnodes = Connector.GetObject(SearchFixedBlocks[0].Index);
                        File.AppendAllText(logfilePath, "Fixed Node Index : " + SearchFixedBlocks[0].Index + Environment.NewLine);
                        File.AppendAllText(logfilePath, "Fixed Node Name : " + testnodes.NodeName + Environment.NewLine);
                        Connector.ShowObject(SearchFixedBlocks[0].Index, true);
                      if (LOADED_BLOCKS.ContainsKey(SearchFixedBlocks[0].NodeName) == true)
                            LOADED_BLOCKS.Remove(SearchFixedBlocks[0].NodeName);
                    }

                    // 언로드
                    Connector.UnloadMultiNode(LOADED_BLOCKS.Values.ToArray());

                    DateTime ed = DateTime.Now;
                    TimeSpan ad = ed - ss;
                    TotalLoad.Add(ad);//

                    // 이동블록 간섭대상으로 추가
                    multiVO.TaskVO.MovingGroupNodes.Clear();
                    multiVO.TaskVO.MovingGroupNodes.Add(GetNodeSearch(multiVO.MOVING_BLOCK)[0]);
                    // 고정블록 간섭대상으로 추가
                    multiVO.TaskVO.FixedGroupNodes.Clear();
                    foreach (string item in multiVO.FIXED_BLOCK)
                    {
                        multiVO.TaskVO.FixedGroupNodes.Add(GetNodeSearch(item)[0]);
                    }
                    multiVO.StartClashTestDate();   // 시작시간 기록
                    taskVO = multiVO.TaskVO;        // 간섭검사 대상 설정
                    ++_nIndex; //_nIndex = i + 1;
                    break;
                }
            }
            File.AppendAllText(logfilePath, "#@#@" + Environment.NewLine);

            Console.WriteLine("#$#$");

            // 검사 수행할 대상이 있는 경우
            if (taskVO != null)
            {
                File.AppendAllText(logfilePath, "#@#@212" + Environment.NewLine);
                string sCheckType = @"정적간섭검사(모델)";
                if (taskVO.CtType == CtType.CtTypeG2G) { sCheckType = @"정적간섭검사(그룹)"; }
                else if (taskVO.CtType == CtType.CtTypeMoving) { sCheckType = @"동적간섭검사"; }
                string time = "";
                Connector.ShowWaitDialogWithText(true, "간섭체크", string.Format("..."));
                string treename = "";
                try
                {
                    bool isfirst = true;
                    treename = taskVO.MovingGroupNodes[0].NodeName;
                    foreach (var a in tree)
                    {
                        //if (a.Contains("/" + TaskList[j].MOVING_BLOCK + "/"))
                        //{
                        if (a.Split('/').Last() == taskVO.MovingGroupNodes[0].NodeName && isfirst)
                        {
                            if (isfirst)
                            {
                                treename = a;
                                isfirst = false;
                            }
                        }
                    }
                    ////
                    Connector.UpdateWaitDialogDescription(string.Format("Total ({0}/{1}) {2} [{3}]", (_nIndex / 3) + 1, TaskList.Count / 3, taskVO.MovingGroupNodes[0].NodeName, sCheckType));

                    time = string.Format("{0}시{1}분{2}초", DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00")
                                    , DateTime.Now.Second.ToString("00"));
                    File.AppendAllText(logfilePath, string.Format("#####################") + Environment.NewLine);
                    File.AppendAllText(logfilePath, string.Format("Total ({0}/{1}) {2}({4}) [{3}]", (_nIndex / 3) + 1, TaskList.Count / 3, taskVO.MovingGroupNodes[0].NodeName, sCheckType, treename) + Environment.NewLine);
                    File.AppendAllText(logfilePath, string.Format("[{0} {1}] 간섭체크 시작 준비", DateTime.Now.ToShortDateString(), time) + Environment.NewLine);
                }
                catch { }
                List<int> fNode = new List<int>();
                foreach (NodeVO item in taskVO.FixedGroupNodes)
                {
                    fNode.Add(item.Index);
                    NodeVO testnodes = Connector.GetObject(item.Index);
                    File.AppendAllText(logfilePath, "fNode Index : " + item.Index + Environment.NewLine);
                    File.AppendAllText(logfilePath, "fNode Name : " + item.NodeName + Environment.NewLine);
                }
                DateTime ss = DateTime.Now;
                Connector.ShowObjects(fNode.ToArray(), true);               // 조회
                Connector.SetObjectsColor(fNode.ToArray(), Color.Blue);     // 색상변경

                List<int> mNode = new List<int>();
                foreach (NodeVO item in taskVO.MovingGroupNodes)
                {
                    mNode.Add(item.Index);
                    NodeVO testnodes = Connector.GetObject(item.Index);
                    File.AppendAllText(logfilePath, "mNode Index : " + item.Index + Environment.NewLine);
                    File.AppendAllText(logfilePath, "mNode Name : " + item.NodeName + Environment.NewLine);
                }
                Connector.ShowObjects(mNode.ToArray(), true);               // 조회
                Connector.SetObjectsColor(mNode.ToArray(), Color.Orange);   // 색상변경
                //20.07.21 J102 모두 제거
                //201202 J101,J102 모두제거 --> 간섭개수가 엄청나오므로 등록에 시간이 걸림
                foreach (var Model in Connector.GetAllObjects())
                {
                    try
                    {
                        if (Model.NodeName == "J102"|| Model.NodeName == "J101")
                        {
                            NodeVO testnodes = Connector.GetObject(Model.Index);
                            File.AppendAllText(logfilePath, "J102orJ101 Index : " + Model.Index + Environment.NewLine);
                            File.AppendAllText(logfilePath, "J102orJ101 Name : " + Model.NodeName + Environment.NewLine);
                            Connector.ShowObject(Model.Index, false);
                        }
                    }
                    catch (Exception ex) { }
                }
                //19.08.28 동적일때 후행모델 제거 Logic 추가
                if (taskVO.CtType == CtType.CtTypeMoving || taskVO.CtType == CtType.CtTypeG2G)
                {
                    foreach (var Model in Connector.GetAllObjects())
                    {
                        try
                        {
                            if (Model.Depth < 1)
                                continue;
                            if (Model.NodeName == "AFTER")
                            {
                                NodeVO testnodes = Connector.GetObject(Model.Index);
                                File.AppendAllText(logfilePath, "AFTER Index : " + Model.Index + Environment.NewLine);
                                File.AppendAllText(logfilePath, "AFTER Name : " + Model.NodeName + Environment.NewLine);
                                Connector.ShowObject(Model.Index, false);
                            }
                        }
                        catch (Exception ex) { }
                    }                  
                }
                //2021-05-17 Lug&Wire 중 _E_ 만 실행

                //일단 Lug&Wire를 모두 제거
                foreach (var Model in Connector.GetAllObjects())
                {
                    if (Model.NodeName == "/LUG")
                    {
                        File.AppendAllText(logfilePath, "LUG Index : " + Model.Index + Environment.NewLine);
                        File.AppendAllText(logfilePath, "LUG Name : " + Model.NodeName + Environment.NewLine);
                        Connector.ShowObject(Model.Index, false);
                    }
                    foreach (var subLug in Connector.GetChildObjects(Model.Index, ChildrenTypes.Children))
                    {
                        foreach (var SubModel in Connector.GetChildObjects(subLug.Index, ChildrenTypes.Children))
                        {
                            try
                            {
                                if (Model.NodeName.Contains(string.Format("_LUG_E__{0}", taskVO.MovingGroupNodes[0].NodeName)))
                                {
                                    File.AppendAllText(logfilePath, "_LUG_E_ Index : " + Model.Index + Environment.NewLine);
                                    File.AppendAllText(logfilePath, "_LUG_E_ Name : " + Model.NodeName + Environment.NewLine);
                                    Connector.ShowObject(SubModel.Index, true);
                                }
                                if (Model.NodeName.Contains(string.Format("_WIRE_E__{0}", taskVO.MovingGroupNodes[0].NodeName)))
                                {
                                    File.AppendAllText(logfilePath, "_WIRE_E_ Index : " + Model.Index + Environment.NewLine);
                                    File.AppendAllText(logfilePath, "_WIRE_E_ Name : " + Model.NodeName + Environment.NewLine);
                                    Connector.ShowObject(SubModel.Index, true);
                                }
                            }
                            catch (Exception ex) { }
                        }
                    }
                }

 
                DateTime ed = DateTime.Now;
                TimeSpan ad = ed - ss;
                TotalLoad.Add(ad);//

                //
                time = string.Format("{0}시{1}분{2}초", DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00")
                        , DateTime.Now.Second.ToString("00"));
                File.AppendAllText(logfilePath, string.Format("[{0} {1}] 모델로드 완료 간섭체크 시작", DateTime.Now.ToShortDateString(), time) + Environment.NewLine);

                if (taskVO.CtType == CtType.CtTypeMoving)
                {
                    bool check = false;
                    foreach (var a in gijangblock)
                    {
                        if (taskVO.MovingGroupNodes[0].NodeName.Contains(a.ToUpper()))
                        {
                            check = true;
                        }
                    }
                    if (check)
                    {
                        taskVO.RangeTolerance = float.Parse(gijang.Text);
                    }
                }
                File.AppendAllText(logfilePath, string.Format("[{0} {1}]  ###1", DateTime.Now.ToShortDateString(), time) + Environment.NewLine);

                //try
                //{
                    bool bStartCheck = Connector.Clash_StartCheck(taskVO); // 간섭검사 수행
                    if (!bStartCheck)
                    {
                        Connector_OnFinishedClashTestEvent(null, null);
                    }
                //}
                //catch (Exception ex)
                //{
                //    File.AppendAllText(logfilePath, string.Format("[{0} {1}]  {2}", DateTime.Now.ToShortDateString(), time, ex.Message) + Environment.NewLine);
                //    Connector_OnFinishedClashTestEvent(null, null);
                //}

            }
            else // 검사가 완료 된 경우
            {
                _checkOver = true;
                Console.WriteLine("***_checkOver:{0}", _checkOver);

                List<string> RESULT = new List<string>();
                Console.WriteLine("TaskList.Count:{0}", TaskList.Count);
                for (int j = 0; j < TaskList.Count; j++)
                {
                    Console.WriteLine("j:{0}", j);
                    if (j != (TaskList.Count - 1)) { continue; }

                    if (TaskList[j].ResultItem.Count == 0) { continue; }
                    if (TaskList[j].TaskVO.CtType != CtType.CtTypeMoving) { continue; }
                    Console.WriteLine("   ~~~~~~~~~~~#");

                    string sCheckTypeName = "간섭체크0";
                    if (TaskList[j].TaskVO.CtType == CtType.CtTypeG2G) { sCheckTypeName = "간섭체크1"; }
                    else if (TaskList[j].TaskVO.CtType == CtType.CtTypeMoving) { sCheckTypeName = "간섭체크2"; }

                    string uppblock = TaskList[j].MOVING_BLOCK;
                    if (allblock)
                    {
                        string block = "";
                        bool isfirst = true;
                        foreach (var a in tree)
                        {
                            //if (a.Contains("/" + TaskList[j].MOVING_BLOCK + "/"))
                            //{
                            if (a.Split('/').Last() == TaskList[j].MOVING_BLOCK && isfirst)
                            {
                                if (isfirst)
                                {
                                    //block = a.Split('/').ToList()[2];
                                    int cas = 0;
                                    foreach (var b in a.Split('/').ToList())
                                    {
                                        if (cas >= 2)
                                        {
                                            block += b + @"\";
                                        }
                                        cas++;

                                    }
                                    isfirst = false;
                                }
                            }
                        }
                        uppblock = block;
                    }
                    _csvFileName = Path.Combine(_selectedSHIPPath, String.Format(@"{0}\{1}_{2}.CSV", uppblock, TaskList[j].MOVING_BLOCK, sCheckTypeName));
                    Console.WriteLine("  _csvFileName:{0}", _csvFileName);
                    //
                    if (_usedCSVFileNames.Contains(_csvFileName) == true) { }
                    else
                    {
                        if (File.Exists(_csvFileName) == true)
                        {
                            try
                            {
                                File.Delete(_csvFileName);
                                Console.WriteLine("   Succeeded, Delete {0}", _csvFileName);
                            }
                            catch
                            {
                                Console.WriteLine("   Failed, Delete {0}", _csvFileName);
                            }
                        }
                        Console.WriteLine("!!!j:{0},_csvFileName:{1}", j, _csvFileName);

                        RESULT.AddRange(TaskList[j].ExportResult99());
                        System.IO.File.WriteAllLines(_csvFileName, RESULT.ToArray(), Encoding.UTF8);
                        _usedCSVFileNames.Add(_csvFileName);

                        if (File.Exists(_csv2csvEXE) == true)
                        {
                            if (File.Exists(_csvFileName) == true)
                            {
                                string runCMDStr = String.Format(@"{0} {1} {2}", _csv2csvEXE, _csvFileName, s_attFileNames);
                                Console.WriteLine("###runCMDStr:{0}", runCMDStr);
                                RunCMDOnly(runCMDStr, true, true);

                                Console.WriteLine("---");
                            }
                        }
                    }
                    Connector.OnFinishedClashTestEvent -= Connector_OnFinishedClashTestEvent;//간섭제외
                    //
                }

                // 전체 수행 시간 계산
                //int nTotalSec = 0;

                //TimeSpan ts = TaskList[0].ClashTestStartDate - TaskList[TaskList.Count-1].ClashTestFinishDate;
                //int retInt = 0;
                //try { retInt = Convert.ToInt32(ts.TotalSeconds); }
                //catch { }
                //nTotalSec = retInt;

                // 다이얼로그 숨기기
                Connector.ShowWaitDialog(false);

                // 모델 상태 초기화
                Connector.RestoreAllObjectColor();

                //모델 거리 재계산 Logic
                //string time = string.Format("{0}시{1}분{2}초", DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00")
                //                , DateTime.Now.Second.ToString("00"));
                //File.AppendAllText(logfilePath, string.Format("[{0} {1}]  모델 거리 재계산 시작", DateTime.Now.ToShortDateString(), time) + Environment.NewLine);
                //reCalculation();
                //time = string.Format("{0}시{1}분{2}초", DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00")
                //                    , DateTime.Now.Second.ToString("00"));
                //File.AppendAllText(logfilePath, string.Format("[{0} {1}]  모델 거리 재계산 끝", DateTime.Now.ToShortDateString(), time) + Environment.NewLine);

                //// if 6003
                //int nHour = nTotalSec / 3600; //1
                //int rest1 = nTotalSec % 3600; //2403
                //int nMin = rest1 / 60; // 40

                // 완료 메시지
                TimeSpan how = DateTime.Now - starttime;
                //if (autorun)
                //{
                //    Connector.Exit(true);
                //}
                //MessageBox.Show(string.Format("탑재간섭검사 완료. 총 소요시간:{0}초.\r\n ({1}시간 {2}분)", how.TotalSeconds, how.Hours, how.Minutes), "VIZZARD", MessageBoxButtons.OK, MessageBoxIcon.Information);

                string time = string.Format("{0}시{1}분{2}초", DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00")
 , DateTime.Now.Second.ToString("00"));
                File.AppendAllText(logfilePath, string.Format("[{0} {1}]간섭체크 완료", DateTime.Now.ToShortDateString(), time) + Environment.NewLine);
                File.AppendAllText(logfilePath, string.Format("[{0} {1}]총 Load 시간 {2}", DateTime.Now.ToShortDateString(), time, TotalLoad.TotalSeconds) + Environment.NewLine);
                //Message뜨면서 강제종료가 안됨 --> 메시지 제거 및 강제 종료
                Connector.KillProcess();
                Connector.Exit(true);
            }
            Console.WriteLine("End, timerClash_Tick");
        }
        public void runcmdforRecal(string rev, string excel)
        {
            ProcessStartInfo proinfo = new ProcessStartInfo();
            Process pro = new Process();

            //실행할 파일명 입력 --cmd
            proinfo.FileName = @"cmd";
            //cmd창 띄우기 -- true(띄우지 않기.)false(띄우기)
            proinfo.CreateNoWindow = true;
            proinfo.UseShellExecute = false;
            //cmd 데이터 받기
            proinfo.RedirectStandardOutput = true;
            //cmd 데이터 보내기
            proinfo.RedirectStandardInput = true;
            //cmd 오류내용 받기
            proinfo.RedirectStandardError = false;
            Console.WriteLine("#1");
            pro.StartInfo = proinfo;
            pro.Start();
            string curvesplit = @"Q:\env\Vitesse\Infoget\autoAP\clashREcheck.exe";

            //cmd에 보낼 명령어를 입력 합니다.
            if (companyName == "IFG")
            {
                curvesplit = @"D:\clashREcheck.exe";
            }
            Console.WriteLine(curvesplit);
            pro.StandardInput.Write(string.Format("{0} {1} {2}", curvesplit, rev, excel)
                + Environment.NewLine);
            pro.StandardInput.Close();
            Console.WriteLine("#2");
            ////결과 값을 리턴 받습니다.
            pro.WaitForExit();
            pro.Close();
            Console.WriteLine("#3");
            //결과 값을 확인 합니다.

            //return resultValue;
        }
        private List<NodeVO> GetNodeSearch(string NodeName, bool islug = false,bool sta = true)
        {
            List<NodeVO> items = new List<NodeVO>();
            if (islug)
            {
                List<NodeVO> children = Connector.FindObject(NodeName, false, true, false, false, false, true);
                List<NodeVO> chil = Connector.GetChildObjects(children[0].Index, ChildrenTypes.Children);
                foreach(var a in chil)
                {
                    //if (a.NodeName.Contains("_L_"))
                    //{ items.Add(a); }
                    //else if (a.NodeName.Contains("_T_"))
                    //{ items.Add(a); }
                    //else
                    //{ items.Add(a); }
                    if(sta)
                    {
                        if (a.NodeName.Contains("_R_"))
                        { items.Add(a); }
                    }
                    else
                    {
                        if (a.NodeName.Contains("_L_"))
                        { items.Add(a); }
                        else if (a.NodeName.Contains("_T_"))
                        { items.Add(a); }
                    }
                }
            }
            else
            {
                items = Connector.FindObject(NodeName, false, true, false, false, false, true);
            }
            return items;
        }

        private void Connector_OnFinishedClashTestEvent(object sender, FinishedClashTestEventArgs e)
        {
            string time = string.Format("{0}시{1}분{2}초", DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00")
, DateTime.Now.Second.ToString("00"));
            File.AppendAllText(logfilePath, string.Format("[{0} {1}] 간섭체크 끝", DateTime.Now.ToShortDateString(), time) + Environment.NewLine);

            Console.WriteLine("Start,  Connector_OnFinishedClashTestEvent");
            // 간섭결과 제외 적용 목록 반환
            List<ClashResultVO> ResultItem = new List<ClashResultVO>();
            try
            {
                ResultItem.AddRange(Connector.GetClashTestExcludingList(e.Result));
            }
            catch
            { }

            ClashTaskMultiVO task = new ClashTaskMultiVO();
            _csvFileName = "";
            List<string> RESULT = new List<string>();

            Console.WriteLine("_iIndex:{0}, TaskList.Count:{1}, ResultItem.Count:{2}***********", _iIndex, TaskList.Count, ResultItem.Count);
            for (int i = 0; i < TaskList.Count; i++)
            {
                task = TaskList[i];

                Console.WriteLine(" ~~~!i:{0}, task.MOVING_BLOCK:{1}, task.TaskVO.CtType.ToString():{2}", i, task.MOVING_BLOCK, task.TaskVO.CtType.ToString());

                if (task.IsTesting == true)
                {
                    Console.WriteLine(" IsTesting@i:{0}, task.MOVING_BLOCK:{1}, task.TaskVO.CtType.ToString():{2}, task.IsCompleted:{3}, task.IsTesting:{4}", i, task.MOVING_BLOCK, task.TaskVO.CtType.ToString(), task.IsCompleted, task.IsTesting);
                    task.IsTesting = false;
                    task.IsCompleted = true;

                    task.ResultItem = ResultItem;
                    task.FinishClashTestDate();

                    if (task.ListItem != null)
                    {
                        try
                        {
                            lvList.Invoke(new EventHandler(delegate
                            {
                                task.ListItem.SubItems[6].Text = string.Format("{0:n0}", ResultItem.Count);
                                task.ListItem.SubItems[7].Text = string.Format("{0:n0}", task.GetElapsedSec());
                            }));
                        }
                        catch { }
                    }

                    if (task.TaskVO.CtType == CtType.CtTypeMoving)
                    {
                        //Connector.IgnoreModelChangedStatus(true);
                        //Connector.OpenDocument(CLASH_MODEL_FILE);
                    }
                }

                if (task.IsCompleted == true)
                {
                    if(Moving.Checked)
                    {
                        if(task.TaskVO.CtType == CtType.CtTypeG2G||
                        task.TaskVO.CtType == CtType.CtTypeSelf)
                        {
                            continue;
                        }
                    }
                    Console.WriteLine(" IsCompleted@i:{0}, task.MOVING_BLOCK:{1}, task.TaskVO.CtType.ToString():{2}, task.IsCompleted:{3}, task.IsTesting:{4}", i, task.MOVING_BLOCK, task.TaskVO.CtType.ToString(), task.IsCompleted, task.IsTesting);

                    string sCheckTypeName = "간섭체크0";
                    if (task.TaskVO.CtType == CtType.CtTypeSelf) { sCheckTypeName = "간섭체크1"; }
                    else if (task.TaskVO.CtType == CtType.CtTypeMoving) { sCheckTypeName = "간섭체크2"; }
                    string uppblock = task.MOVING_BLOCK;
                    if (allblock)
                    {
                        string block = "";
                        bool isfirst = true;
                        foreach (var a in tree)
                        {
                            if (a.Split('/').Last() == task.MOVING_BLOCK && isfirst)
                            {
                                if (isfirst)
                                {
                                    //block = a.Split('/').ToList()[2];
                                    int cas = 0;
                                    bool ccc = false;
                                    foreach (var b in a.Split('/').ToList())
                                    {
                                        if(b==uppblock)
                                        {
                                            block += b;
                                            ccc = true;
                                        }
                                        if(ccc)
                                        {
                                            continue;
                                        }
                                        if (cas >= 2)
                                        {
                                            block += b + @"\";
                                        }
                                        cas++;

                                    }
                                    isfirst = false;
                                }
                            }
                        }
                        uppblock = block;
                    }
                    _csvFileName = Path.Combine(_selectedSHIPPath, String.Format(@"{0}\{1}_{2}.CSV", uppblock, task.MOVING_BLOCK, sCheckTypeName));
                    string imagePath= Path.Combine(_selectedSHIPPath, String.Format(@"{0}\{1}", uppblock, sCheckTypeName));
                    Console.WriteLine("  _csvFileName:{0}", _csvFileName);

                    if (_usedCSVFileNames.Contains(_csvFileName) == true) { }
                    else
                    {
                        if (File.Exists(_csvFileName) == true)
                        {
                            try
                            {
                                File.Delete(_csvFileName);
                                Console.WriteLine("   Succeeded, Delete {0}", _csvFileName);
                            }
                            catch
                            {
                                Console.WriteLine("   Failed, Delete {0}", _csvFileName);
                            }
                        }
                        Console.WriteLine("~~~_csvFileName:{0}", _csvFileName);

                        try
                        {
                            RESULT.AddRange(task.ExportResult99(imagePath, imageCheck.CheckState));
                        }
                        catch(Exception ex) { MessageBox.Show(ex.Message); }
                        Console.WriteLine("A");
                        ///JMS 2019-09-18일 수정 마린소프트 건
                        if (marinbool)
                        {
                            if (dirbool)
                            {
                                if (_csvFileName.Contains("간섭체크1"))
                                {
                                    string filename = task.MOVING_BLOCK.Replace(".REV", "");
                                    string finalpath = Path.Combine(filepathM, String.Format(@"{0}.CSV", filename));
                                    System.IO.File.WriteAllLines(finalpath, RESULT.ToArray(), Encoding.UTF8);
                                }
                            }
                            else
                            {
                                if (_csvFileName.Contains("간섭체크1"))
                                {
                                    DirectoryInfo mar = new DirectoryInfo(filepathM);
                                    if (!mar.Exists)
                                    {
                                        mar.Create();
                                    }

                                    if (RESULT.Count != 0)
                                    {
                                        string finalpath = Path.Combine(filepathM, filenameM + ".csv");
                                        System.IO.File.WriteAllLines(finalpath, RESULT.ToArray(), Encoding.UTF8);
                                    }
                                    else
                                    {
                                        string finalpath = Path.Combine(filepathM, filenameM + ".err");
                                        System.IO.File.WriteAllLines(finalpath, RESULT.ToArray());
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (RESULT.Count != 0)
                            {
                                if (_csvFileName.Contains("간섭체크0"))
                                {
                                    string rename = _csvFileName.Replace("간섭체크0", "간섭체크1");;
                                    var lineCount = File.ReadLines(rename).Count();
                                    List<string> countre = new List<string>();
                                    foreach(var a in RESULT)
                                    {
                                        lineCount++;
                                        string adddata = "";
                                        List<string> split = a.Split('∫').ToList();
                                        split[0] = lineCount.ToString();
                                        bool isfirst = true;
                                        foreach(var b in split)
                                        {
                                            if (isfirst)
                                            {
                                                adddata += b;
                                                isfirst = false;
                                            }
                                            else
                                            {
                                                adddata += "∫" + b;
                                            }
                                        }
                                        countre.Add(adddata);
                                    }
                                    System.IO.File.AppendAllLines(rename, countre.ToArray(), Encoding.UTF8);
                                }
                                else
                                {
                                    System.IO.File.WriteAllLines(_csvFileName, RESULT.ToArray(), Encoding.UTF8);
                                }
                            }
                            else
                            {
                                System.IO.File.WriteAllLines(_csvFileName, RESULT.ToArray());
                            }
                        }
                        _usedCSVFileNames.Add(_csvFileName);

                        if (File.Exists(_csv2csvEXE) == true)
                        {
                            /*if (File.Exists(_csvFileName) == true)
                            {
                                string runCMDStr = String.Format(@"{0} {1} {2}", _csv2csvEXE, _csvFileName, s_attFileNames);
                                Console.WriteLine("###runCMDStr:{0}", runCMDStr);
                                RunCMDOnly(runCMDStr, true, true);

                                Console.WriteLine("---");
                            }*/
                            try
                            {
                                File.Delete(_csv2csvEXE);  //temporary
                                Console.WriteLine("   Succeeded, Delete {0}", _csv2csvEXE);
                            }
                            catch
                            {
                                Console.WriteLine("   Failed, Delete {0}", _csv2csvEXE);
                            }
                        }
                    }

                    //timerClash.Enabled = true;

                    //continue;
                }
            }
            ++_iIndex;
            
            Console.WriteLine("~~~!!!@@@");

            time = string.Format("{0}시{1}분{2}초", DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00")
, DateTime.Now.Second.ToString("00"));
            File.AppendAllText(logfilePath, string.Format("[{0} {1}]결과 파일 생성 완료", DateTime.Now.ToShortDateString(), time) + Environment.NewLine);
            File.AppendAllText(logfilePath, string.Format("[{0} {1}]결과 파일 경로 : {2}", DateTime.Now.ToShortDateString(), time, _csvFileName) + Environment.NewLine);
            File.AppendAllText(logfilePath, string.Format("#####################") + Environment.NewLine);
            timerClash.Enabled = true;

            Console.WriteLine("End,  Connector_OnFinishedClashTestEvent");
        }

        private bool createCSV2CSV(string fileName)
        {
            Console.WriteLine("Start, createCSV2CSV");
            bool bSuccess = false;

            ArrayList writeLines = new ArrayList();
            string fileNameNEW = String.Format("{0}", fileName.Replace("간섭체크", "간섭체크필터링"));
            Console.WriteLine("fileNameNEW:{0}", fileNameNEW);
            try
            {
                FileStream fileStreamOutput = new FileStream(fileNameNEW, FileMode.Create);
                fileStreamOutput.Seek(0, SeekOrigin.Begin);

                ArrayList usedModel1Model2 = new ArrayList();

                //
                string[] readLines = File.ReadAllLines(fileName, Encoding.Default);
                Console.WriteLine("readLines.Length:{0}", readLines.Length);
                foreach (string readLine in readLines)
                {

                    string[] tt = readLine.Split(',');
                    string model1 = String.Format("{0}", tt[3]);
                    if (model1.Contains(" /") == true) { model1 = String.Format("{0}", model1.Substring(model1.LastIndexOf(" /")).Trim()); }
                    string model2 = String.Format("{0}", tt[4]);
                    if (model2.Contains(" /") == true) { model2 = String.Format("{0}", model2.Substring(model2.LastIndexOf(" /")).Trim()); }
                    string Model1Model2 = String.Format("{0}#{1}", model1, model2);
                    if (usedModel1Model2.Contains(Model1Model2) == true) { continue; }
                    else { usedModel1Model2.Add(Model1Model2); }

                    string writeLine = String.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24}",
                        tt[0], tt[1], tt[2], model1, model2, tt[5], tt[6], tt[7], tt[8], tt[9], tt[10], tt[11], tt[12], tt[13], tt[14], tt[15], tt[16], tt[17], tt[18], tt[19], tt[20], tt[21], tt[22], tt[23], tt[24]);

                    byte[] info;

                    info = System.Text.Encoding.Default.GetBytes(writeLine + "\r\n"); //ANSI
                    //info = new UTF8Encoding(true).GetBytes(writeLine + "\r\n"); //UTF-8
                    fileStreamOutput.Write(info, 0, info.Length);

                }
                //

                fileStreamOutput.Flush();
                fileStreamOutput.Close();

                bSuccess = true;
            }
            catch
            {
                //
            }

            return bSuccess;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Console.WriteLine("Start, button1_Click");
                //로그파일생성
                string time2 = string.Format("{0}시{1}분{2}초", DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00"), DateTime.Now.Second.ToString("00"));
                string path2 = @"C:\temp";
                logfilePath = Path.Combine(path2, string.Format("{0}({1}_정동적실행).log", DateTime.Today.ToShortDateString(), time2));
                if (logFileName != "")
                {
                    logfilePath = logFileName;//arg에서 받은 Log파일 경로 설정
                }

                if (!(new DirectoryInfo(path2)).Exists)
                {
                    Console.WriteLine(string.Format("LOG 경로를 확인하세요 : {0}", path2));
                }
                Console.WriteLine("Log파일 이름 부여");

                StreamWriter log = new StreamWriter(logfilePath, true);
                starttime = DateTime.Now;
                log.WriteLine(string.Format("[{0} {1}]Log File", DateTime.Now.ToShortDateString(), time2));
                //log.WriteLine(string.Format("{0}_{1} 실행", DateTime.Now.ToShortDateString(),time));
                log.Close();
                if(autorun)
                {
                    foreach (string selblock in selectblock)
                    {
                        File.AppendAllText(logfilePath,string.Format("Selected : {0}",selblock));
                    }
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
                return;
            }

            _nIndex = 0;
            _checkOver = false;

            if (watListView1.SelectedItems.Count == 0)
            {
                MessageBox.Show(String.Format(@"Listview 선택은 0보다 커야 합니다.(현재 선택:{0})", watListView1.SelectedItems.Count));
                return;
            }

            _usedCSVFileNames.Clear();

            //
            bool bContinue = true;
            getSystemInfo();
            double limitMemory = 7.9 * 1024;
            if (_dMemory < limitMemory) { bContinue = false; }
            if (_computerName == "D04252") { bContinue = true; } //HMD 교육장.
            if (_computerName == "D06047") { bContinue = true; } //원격 PC
            if (_computerName == "PC04335D") { bContinue = true; } //HMD 교육장
            if (_computerName == "DESKTOP-VS2M175") { bContinue = true; } //HMD 원격PC-박성준
            if (_computerName == "HMD") { bContinue = true; } //HMD 원격PC-권오욱
            if (bContinue == false)
            {
                MessageBox.Show(String.Format(@"[Error] PC 전체 메모리는 {0} MB보다 커야 합니다.(현재 메모리:{1} MB)", limitMemory, _dMemory));
                return;
            }
            //

            tabControl1.SelectedTab = tabPage2;

            _iIndex = 0;
            //Console.WriteLine("----------------------------TEMP 모델 제거");
            //REVreCreate();
            //Connector.IgnoreModelChangedStatus(true);
            //Connector.CloseDocument();

            //첫번째 탑재블록 0,2 지우기 Logic
            string block = watListView1.Items[0].SubItems[0].Text;
            string bb1 = Path.Combine(_selectedSHIPPath, block, block+"_간섭체크0.CSV");
            string bb2 = Path.Combine(_selectedSHIPPath, block, block + "_간섭체크2.CSV");
            FileInfo ff = new FileInfo(bb1);
            if(ff.Exists)
            {
                ff.Delete();
            }
            FileInfo ff2= new FileInfo(bb2);
            if (ff2.Exists)
            {
                ff2.Delete();
            }
            foreach(var ex in exceptBlock)
            {
                string ex2 = Path.Combine(_selectedSHIPPath, ex, ex + "_간섭체크2.CSV");
                FileInfo ffb = new FileInfo(ex2);
                if (ffb.Exists)
                {
                    ffb.Delete();
                }
            }

            Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~0");
            btnLoadModel_Click(null, null);

            Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~1");
            btnShowModel_Click(null, null);

            Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Pre Check");

            //try
            //{
            //    bool keepgoing = preCheck();

            //    if (!keepgoing)
            //    {
            //        return;
            //    }
            //}
            //catch
            //{ }
            Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~2");
            btnClashTest_Click(null, null);

            Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~3");
            btnStartClashTest_Click(null, null);

            Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~4");

            Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~5");

            Console.WriteLine("End, button1_Click");
            //if (autorun)
            //{
            //    Connector.Exit(true);
            //}
        }   
        private void cbBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            allblock = false;
            Console.WriteLine("\n\n\n\n\nStart, cbBox1_SelectedIndexChanged");
            //
            if (keyin)
            {
                if (!cbBox1.Items.Contains(cbBox1.Text))
                {
                    return;
                }
                cbBox1.SelectedItem = cbBox1.Text;
            }

            tabControl1.SelectedTab = tabPage1;

            watListView1.Items.Clear();
            listView2.Items.Clear();
            if (cbBox1.SelectedItem.ToString() == "") { return; }

            //Console.WriteLine("_originPath:{0}", _originPath);
            _selectedSHIPPath = Path.Combine(_originPath, cbBox1.SelectedItem.ToString());
            _erectPath = Path.Combine(_originPath, "ERECT", cbBox1.SelectedItem.ToString());
            DirectoryInfo CC = new DirectoryInfo(_erectPath);
            if(!CC.Exists)
            {
                CC.Create();
            }

            //
            _shipNo = _selectedSHIPPath.Substring(_selectedSHIPPath.LastIndexOf(@"\") + 1);
            Console.WriteLine("_shipNo:{0}", _shipNo);
            proj = "AM121KBS";
            if (companyName == "HMD")
            {
                string projectDir = String.Format(@"Q:\project\{0}", _shipNo);
                Console.WriteLine("projectDir:{0}", projectDir);

                if (Directory.Exists(projectDir) == true)
                {
                    Console.WriteLine("Exist, projectDir:{0}", projectDir);

                    DirectoryInfo dir = new DirectoryInfo(projectDir);
                    DirectoryInfo[] subFiles = dir.GetDirectories();
                    foreach (DirectoryInfo ff in subFiles)
                    {
                        string checkString = ff.Name.ToString().ToUpper().Trim();
                        Console.WriteLine("checkString:{0}", checkString);

                        if (checkString.EndsWith("000") == true)
                        {
                            proj = checkString.Substring(0, checkString.LastIndexOf("000"));
                            break;
                        }
                    }
                }
            }
            Console.WriteLine("proj:{0}", proj);

            //
            _tempDir2 = Path.Combine(_tempDir, String.Format("{0}_{1}_{2}", _shipNo, myProc.Id, am12USERNAME));
            if (Directory.Exists(_tempDir2) == false)
            {
                try { Directory.CreateDirectory(_tempDir2); }
                catch { Console.WriteLine("Failed, Create _tempDir2:{0}", _tempDir2); }
            }
            //

            Console.WriteLine("_selectedSHIPPath:{0}, _shipNo:{1}", _selectedSHIPPath, _shipNo);

            if (String.IsNullOrEmpty(_selectedSHIPPath) == true) { return; }

            dic_treenode2Path.Clear();
            _treeAllFile = Path.Combine(_selectedSHIPPath, @"Z99_TREE_ALL.TXT");
            erectBlocks.Clear();
            if (File.Exists(_treeAllFile) == true)
            {
                string[] readLines = File.ReadAllLines(_treeAllFile, Encoding.Default);
                tree = readLines.ToList();
                foreach (string readLine in readLines)
                {

                    if (readLine == "") { continue; }

                    string readLineReplace = Path.Combine(_originPath, String.Format("{0}", readLine.Replace(@"/", @"\").Substring(1)));
                    //Console.WriteLine("readLineReplace:{0}", readLineReplace);

                    string nodeName = "";
                    try { nodeName = readLine.Substring(readLine.LastIndexOf(@"/") + 1).Trim(); }
                    catch { }

                    try
                    {
                        List<string> splitdata = readLine.Split('/').ToList();
                        if(splitdata.Count>=3)
                        {
                            if (!erectBlocks.Contains(splitdata[2]))
                            {
                                erectBlocks.Add(splitdata[2]);
                            }
                        }
                    }
                    catch { }

                    if (dic_treenode2Path.ContainsKey(nodeName) == true) { }
                    else { dic_treenode2Path[nodeName] = readLineReplace; }
                }
            }
            else
            {
                MessageBox.Show(string.Format("{0}의 탑재일자 파일이 없습니다 탑재일자 파일 :{1}", _shipNo, _treeAllFile));
            }
            Console.WriteLine("dic_treenode2Path.Count:{0}", dic_treenode2Path.Count);

            //
            dic_Block2Date.Clear();
            List<string> erectlist = null;
            _erectDateFile = Path.Combine(_selectedSHIPPath, @"Z99_ERECT_ALL.TXT");
            if (File.Exists(_erectDateFile) == true)
            {
                erectlist = new List<string>();
                erectlist = ReadTxt(_erectDateFile);
            }
            if (erectlist != null)
            {
                foreach (string erectl in erectlist)
                {
                    string erectB = erectl.Split(',')[0].Trim();
                    string erectD = erectl.Split(',')[1].Trim();
                    if (!dic_Block2Date.ContainsKey(erectB))
                    {
                        dic_Block2Date[erectB] = erectD;
                    }
                }
            }
            Console.WriteLine("dic_Block2Date.Count:{0}", dic_Block2Date.Count);
            //

            //
            dic_Block2No.Clear(); dic_BlockDir.Clear(); dic_BlockLength.Clear();
            List<string> erectSavedlist = null;
            _erectSavedFile = Path.Combine(_erectPath, @"Z98_ERECT_SAVED.TXT");
            if (File.Exists(_erectSavedFile) == true)
            {
                erectSavedlist = new List<string>();
                erectSavedlist = ReadTxt(_erectSavedFile);
            }
            if (erectSavedlist != null)
            {
                foreach (string erectSaved in erectSavedlist)
                {
                    if (erectSaved.StartsWith("#") == true)
                    {
                        try { _savedUser = erectSaved.Split(',')[1].Trim(); }
                        catch { }

                        continue;
                    }

                    string erectB = erectSaved.Split(',')[0].Trim();
                    string erectD = erectSaved.Split(',')[1].Trim();
                    if (erectSaved.Split(',').ToList().Count >= 3)
                    {
                        string erectDir = erectSaved.Split(',')[2].Trim();
                        if (!dic_BlockDir.ContainsKey(erectB))
                        {
                            dic_BlockDir.Add(erectB, erectDir);
                        }
                    }
                    if (erectSaved.Split(',').ToList().Count >= 4)
                    {
                        string erectLength = erectSaved.Split(',')[3].Trim();
                        if (!dic_BlockLength.ContainsKey(erectB))
                        {
                            dic_BlockLength.Add(erectB, erectLength);
                        }
                    }
                    if (!dic_Block2No.ContainsKey(erectB))
                    {
                        dic_Block2No.Add(erectB,erectD);
                    }
                }
            }
            Console.WriteLine("dic_Block2No.Count:{0}", dic_Block2No.Count);
            //

            label11.Text = String.Format("탑재일자 수정 ({0})", _savedUser);
            ListViewUpdates();

            //
            Console.WriteLine("End, cbBox1_SelectedIndexChanged");
        }

        void ListViewUpdates(List<string> ablock = null)
        {
            Console.WriteLine("\nStart, ListViewUpdates");

            watListView1.Items.Clear();
            watListView1.View = View.Details;
            watListView1.GridLines = true;
            watListView1.FullRowSelect = true;
            watListView1.ComboToList = true;

            watListView2.Items.Clear();
            watListView2.View = View.Details;
            watListView2.GridLines = true;
            watListView2.FullRowSelect = true;
            watListView2.ComboToList = true;

            listView2.Items.Clear();
            listView2.View = View.Details;
            listView2.GridLines = true;
            listView2.FullRowSelect = true;

            DirectoryInfo dir = new DirectoryInfo(_selectedSHIPPath);
            List<DirectoryInfo> subFiles = new List<DirectoryInfo>();
            foreach (var a in dir.GetDirectories())
            {
                if (erectBlocks.Contains(a.Name))
                {
                    subFiles.Add(a);
                }
            }
            //
            int i = 0;
            if (ablock != null)
            {
                foreach (string ff in ablock)
                {
                    string dirName = ff;
                    string selectedItem = ff;
                    string revFile = Path.Combine(_selectedSHIPPath, String.Format(@"{0}\{1}.rev", selectedItem, selectedItem));//Moon 수정 위치(Rev -> viz)//re
                    string vizFile = Path.Combine(_selectedSHIPPath, String.Format(@"{0}\{1}_동적.viz", selectedItem, selectedItem));//Moon 수정 위치(Rev -> viz)//re
                    string csvFile = Path.Combine(_selectedSHIPPath, String.Format(@"{0}\{1}_간섭체크1.CSV", selectedItem, selectedItem));
                    //////Console.WriteLine("dirName:{0}, selectedItem:{1}, revFile:{2}, csvFile:{3}", dirName, selectedItem, revFile, csvFile);

                    if (selectedItem == "HULLPARTDATA") { continue; }
                    if (selectedItem == "ERECT") { continue; }
                    if (selectedItem == "BOUNDARY_INFO") { continue; }
                    if (selectedItem.IndexOf(".") >= 0) { continue; }
                    if (selectedItem.IndexOf("후행") >= 0 || selectedItem.IndexOf("AFTER") >= 0)
                    {
                        if (selectedItem.IndexOf("MAIN_ENGINE") >= 0) { }
                        else if (selectedItem.IndexOf("BOILER") >= 0) { }
                        else { continue; }
                    }
                    if (selectedItem.IndexOf("모델") >= 0 || selectedItem.IndexOf("MODEL") >= 0) { continue; }

                    if (selectedItem != "LUG_WIRE")
                    {
                        bool color = true;
                        string[] strArrs = new string[maxCols];
                        for (int j = 0; j < maxCols; j++)
                        {
                            if (j == 0)
                            {
                                //호선하위노드
                                strArrs[j] = String.Format("{0}", selectedItem);
                            }
                            else if (j == 1)
                            {
                                //REV 생성일자
                                string extractedTime = "";
                                if (File.Exists(revFile) == true)
                                    extractedTime = File.GetLastWriteTime(revFile).ToString();
                                else
                                    extractedTime = "REV 미생성"; //Aveva Marine 12 상의 Assembly Planing에 해당 탑재블록 없음

                                strArrs[j] = String.Format("{0}", extractedTime);
                            }
                            else if (j == 2)
                            {
                                //VIZ 생성일자
                                string extractedTime = "";
                                if (File.Exists(vizFile) == true)
                                    extractedTime = File.GetLastWriteTime(vizFile).ToString();
                                else
                                    extractedTime = "VIZ 미생성"; //Aveva Marine 12 상의 Assembly Planing에 해당 탑재블록 없음

                                strArrs[j] = String.Format("{0}", extractedTime);
                            }
                            else if (j == 3)
                            {
                                //기본순서
                                strArrs[j] = String.Format("{0:D4}", i++);
                            }
                            else if (j == 4)
                            {
                                //REV 탑재일자
                                if (dic_Block2Date.ContainsKey(selectedItem))
                                {
                                    strArrs[j] = dic_Block2Date[selectedItem];
                                }
                                else { strArrs[j] = "Empty"; }
                            }
                            else if (j == 5)
                            {
                                //CSV(간섭결과) 생성일자
                                string extractedCSVTime = "";
                                if (File.Exists(csvFile) == true)
                                    extractedCSVTime = File.GetLastWriteTime(csvFile).ToString();
                                else
                                {
                                    extractedCSVTime = "간섭체크 미실행";
                                    color = false;
                                }

                                strArrs[j] = String.Format("{0}", extractedCSVTime);
                            }
                            else if (j == 6)
                            {
                                //탑재순서
                                strArrs[j] = String.Format("000.0");
                                if (dic_Block2No.ContainsKey(selectedItem))
                                {
                                    strArrs[j] = dic_Block2No[selectedItem];
                                }
                                else
                                {
                                }
                            }
                            else if (j == 7)
                            {
                                string direction = "";
                                try
                                {
                                    if (dic_BlockDir.Keys.Contains(selectedItem))
                                    {
                                        direction = dic_BlockDir[selectedItem];
                                    }
                                    else
                                    {
                                        direction = "Z";
                                    }
                                }
                                catch { MessageBox.Show(selectedItem.ToString()); }
                                strArrs[j] = direction;
                            }
                            else if (j == 8)
                            {
                                string Length = "";
                                try
                                {
                                    if (dic_BlockLength.Keys.Contains(selectedItem))
                                    {
                                        Length = dic_BlockLength[selectedItem];
                                    }
                                    else
                                    {
                                        Length = "200mm";
                                    }
                                }
                                catch { MessageBox.Show(selectedItem.ToString()); }
                                strArrs[j] = Length;
                            }
                            else
                            {
                                strArrs[j] = "";
                            }
                        }
                        ListViewItem strArr = new ListViewItem(strArrs);
                        if (color)
                            strArr.ForeColor = Color.Blue;
                        else
                            strArr.ForeColor = Color.Red;
                        watListView1.Items.Add(strArr);
                    }
                }
            }
            else
            {
                //
                i = 0;
                
                foreach (DirectoryInfo ff in subFiles)
                {
                    bool isexcep = false;
                    string dirName = ff.FullName.ToString().ToUpper().Trim();
                    string selectedItem = ff.Name.ToUpper().Trim();
                    string revFile = Path.Combine(_selectedSHIPPath, String.Format(@"{0}\{1}.rev", selectedItem, selectedItem));//Moon 수정 위치(Rev -> viz)//re
                    string vizFile = Path.Combine(_selectedSHIPPath, String.Format(@"{0}\{1}_동적.viz", selectedItem, selectedItem));//Moon 수정 위치(Rev -> viz)//re
                    string csvFile = Path.Combine(_selectedSHIPPath, String.Format(@"{0}\{1}_간섭체크1.CSV", selectedItem, selectedItem));
                    //////Console.WriteLine("dirName:{0}, selectedItem:{1}, revFile:{2}, csvFile:{3}", dirName, selectedItem, revFile, csvFile);
                    #region Lug
                    if (selectedItem == "LUG_WIRE")
                    {

                        if (selectedItem == "HULLPARTDATA") { continue; }
                        if (selectedItem == "ERECT") { continue; }
                        if (selectedItem == "BOUNDARY_INFO") { continue; }
                        if (selectedItem.IndexOf(".") >= 0) { continue; }
                        if (selectedItem.IndexOf("후행") >= 0 || selectedItem.IndexOf("AFTER") >= 0)
                        {
                            if (selectedItem.IndexOf("MAIN_ENGINE") >= 0) { }
                            else if (selectedItem.IndexOf("BOILER") >= 0) { }
                            else { continue; }
                        }
                        if (selectedItem.IndexOf("모델") >= 0 || selectedItem.IndexOf("MODEL") >= 0) { continue; }

                        DirectoryInfo dirLUG = new DirectoryInfo(dirName);
                        DirectoryInfo[] subLUGFiles = dirLUG.GetDirectories();

                        foreach (DirectoryInfo subLUG in subLUGFiles)
                        {
                            string subName = subLUG.FullName.ToString().ToUpper().Trim();
                            string subItem = subLUG.Name.ToUpper().Trim();
                            string lastName = String.Format("{0}", subItem.Substring(subItem.LastIndexOf("_") + 1));

                            //string LugFile = Path.Combine(_selectedSHIPPath, String.Format(@"{0}\{1}_간섭체크11.CSV", lastName, lastName));
                            //try
                            //{
                            //    LugFile = Path.Combine(dic_treenode2Path[lastName], lastName + "_간섭체크11.CSV");
                            //}
                            //상위 블록 가져오기
                            string block = "";
                            bool isfirst = true;
                            foreach (var a in tree)
                            {
                                if (a.Contains("/" + lastName + "/"))
                                {
                                    if (isfirst)
                                    {
                                        block = a.Split('/').ToList()[2];
                                        isfirst = false;
                                    }
                                }
                            }
                            string LugFile = "";
                            if (block != "")
                            {
                                LugFile = Path.Combine(dic_treenode2Path[block], lastName + "_간섭체크11.CSV");
                            }


                            Console.WriteLine("lastName:{0}, subName:{1}, subItem:{2}", lastName, subName, subItem);

                            string[] str2Arrs = new string[7]; //

                            if (dic_treenode2Path.ContainsKey(lastName) == true)
                            {
                                string pathDic = dic_treenode2Path[lastName];
                                string vizxml = Path.Combine(pathDic, String.Format("{0}.vizxml", lastName));
                                bool bExist_vizxml = File.Exists(vizxml);
                                //Console.WriteLine(" vizxml:{0}, pathDic:{1}, bExist_vizxml:{2}", vizxml, pathDic, bExist_vizxml);

                                if (bExist_vizxml == false) {; }//continue; }

                                string extractedLugWireVIZXMLTime = "";
                                try { extractedLugWireVIZXMLTime = File.GetLastWriteTime(vizxml).ToString(); }
                                catch { }

                                string extractedLugWireCSVFile = Path.Combine(pathDic, String.Format(@"{0}_간섭체크101.CSV", lastName));
                                string extractedLugWireCSVTime = "";
                                if (File.Exists(extractedLugWireCSVFile) == true)
                                {
                                    try { extractedLugWireCSVTime = File.GetLastWriteTime(extractedLugWireCSVFile).ToString(); }
                                    catch { }
                                }
                                Console.WriteLine(" extractedLugWireCSVFile:{0}, extractedLugWireCSVTime:{1}", extractedLugWireCSVFile, extractedLugWireCSVTime);

                                string extractedCSVTime = ""; bool color = true;
                                if (File.Exists(LugFile) == true)
                                    extractedCSVTime = File.GetLastWriteTime(LugFile).ToString();
                                else
                                {
                                    extractedCSVTime = "간섭체크 미실행";
                                    color = false;
                                }

                                str2Arrs[0] = String.Format("{0}", lastName);
                                str2Arrs[1] = extractedLugWireVIZXMLTime;
                                str2Arrs[2] =
                                str2Arrs[3] = "";
                                str2Arrs[4] = String.Format("{0}", extractedCSVTime);//extractedLugWireCSVTime;
                                str2Arrs[5] = "";
                                str2Arrs[6] = String.Format("{0}", vizxml);

                                ListViewItem str2Arr = new ListViewItem(str2Arrs);
                                if (color)
                                    str2Arr.ForeColor = Color.Blue;
                                else
                                    str2Arr.ForeColor = Color.Red;
                                listView2.Items.Add(str2Arr);
                            }
                            else { continue; }
                        }
                    }
                    #endregion
                    else
                    {
                        if (selectedItem == "HULLPARTDATA") { isexcep = true; }
                        if (selectedItem == "2NDPE") { isexcep = true; }
                        if (selectedItem == "ERECT") { isexcep = true;  }
                        if (selectedItem == "AFTE") { isexcep = true;  }
                        if (selectedItem == "BOUNDARY_INFO") { isexcep = true; ; }
                        if (selectedItem.IndexOf(".") >= 0) { isexcep = true; ; }
                        if (selectedItem.IndexOf("후행") >= 0 || selectedItem.IndexOf("AFTER") >= 0)
                        {
                            if (selectedItem.IndexOf("MAIN_ENGINE") >= 0) { }
                            else if (selectedItem.IndexOf("BOILER") >= 0) { }
                            else { isexcep = true; }
                        }
                        if (selectedItem.IndexOf("모델") >= 0 || selectedItem.IndexOf("MODEL") >= 0) { isexcep = true; ; }

                        bool color = true;
                        string[] strArrs = new string[maxCols];
                        for (int j = 0; j < maxCols; j++)
                        {
                            if (j == 0)
                            {
                                //호선하위노드
                                strArrs[j] = String.Format("{0}", selectedItem);
                            }
                            else if (j == 1)
                            {
                                //VIZ 생성일자
                                string extractedTime = "";
                                if (File.Exists(revFile) == true)
                                    extractedTime = File.GetLastWriteTime(revFile).ToString();
                                else
                                    extractedTime = "REV 미추출"; //Aveva Marine 12 상의 Assembly Planing에 해당 탑재블록 없음

                                strArrs[j] = String.Format("{0}", extractedTime);
                            }
                            else if (j == 2)
                            {
                                //VIZ 생성일자
                                string extractedTime = "";
                                if (File.Exists(vizFile) == true)
                                    if (companyName == "IFG")
                                    {
                                        extractedTime = "IFG VIZ";
                                    }
                                    else
                                    {
                                        extractedTime = File.GetLastWriteTime(vizFile).ToString();
                                    }
                                else
                                    extractedTime = "VIZ 미추출"; //Aveva Marine 12 상의 Assembly Planing에 해당 탑재블록 없음

                                strArrs[j] = String.Format("{0}", extractedTime);
                            }
                            else if (j == 3)
                            {
                                //기본순서
                                strArrs[j] = String.Format("{0:D4}", i++);
                            }
                            else if (j == 4)
                            {
                                //탑재일자
                                if (dic_Block2Date.ContainsKey(selectedItem))
                                {
                                    strArrs[j] = dic_Block2Date[selectedItem];
                                }
                                else { strArrs[j] = "Empty"; }
                            }
                            else if (j == 5)
                            {
                                //CSV(간섭결과) 생성일자
                                string extractedCSVTime = "";
                                if (File.Exists(csvFile) == true)
                                    extractedCSVTime = File.GetLastWriteTime(csvFile).ToString();
                                else
                                {
                                    extractedCSVTime = "간섭체크 미실행";
                                    color = false;
                                }

                                strArrs[j] = String.Format("{0}", extractedCSVTime);
                            }
                            else if (j == 6)
                            {
                                strArrs[j] = String.Format("000.0");
                                if (dic_Block2No.ContainsKey(selectedItem))
                                {
                                    strArrs[j] = dic_Block2No[selectedItem];
                                }
                                else
                                {
                                }
                            }
                            else if (j == 7)
                            {
                                string direction = "";
                                try
                                {
                                    if (dic_BlockDir.Keys.Contains(selectedItem))
                                    {
                                        direction = dic_BlockDir[selectedItem];
                                    }
                                    else
                                    {
                                        direction = "Z";
                                    }
                                }
                                catch { MessageBox.Show(selectedItem.ToString()); }
                                strArrs[j] = direction;
                            }
                            else if (j == 8)
                            {
                                string Length = "";
                                try
                                {
                                    if (dic_BlockLength.Keys.Contains(selectedItem))
                                    {
                                        Length = dic_BlockLength[selectedItem];
                                    }
                                    else
                                    {
                                        Length = "200mm";
                                    }
                                }
                                catch { MessageBox.Show(selectedItem.ToString()); }
                                strArrs[j] = Length;
                            }
                            else
                            {
                                strArrs[j] = "";
                            }
                        }
                        ListViewItem strArr = new ListViewItem(strArrs);
                        if (color)
                            strArr.ForeColor = Color.Blue;
                        else
                            strArr.ForeColor = Color.Red;

                        if (isexcep)
                        {
                            watListView2.Items.Add(strArr);
                        }
                        else
                        {
                            watListView1.Items.Add(strArr);
                        }
                    }
                }
            }
            int iSortIndex = 4;
            if (dic_Block2No.Count > 0)
            {
                iSortIndex = 6;
            }
            Console.WriteLine("iSortIndex:{0}", iSortIndex);
            watListView1.ListViewItemSorter = new ListViewItemComparer(iSortIndex, "asc");
            watListView1.ListViewItemSorter = null;
            try
            {
                if (iSortIndex == 4)
                {
                    ListViewItem ME = null;
                    int index = 0; int moveindex = 0;
                    foreach (ListViewItem a in watListView1.Items)
                    {
                        if (a.SubItems[0].Text.Contains("MAIN_ENGINE"))
                        {
                            index = a.Index;
                            ME = a;
                        }
                        if (a.SubItems[0].Text.Contains("5E41") || a.SubItems[0].Text.Contains("2E41"))
                        {
                            moveindex = a.Index;
                        }
                    }
                    ListViewItem oItem = (ListViewItem)watListView1.Items[index].Clone();
                    if (index != 0 && moveindex != 0 && ME != null)
                    {
                        MoveItem(index, moveindex, oItem);
                    }

                    //한사이클
                    ME = null;
                    index = 0; moveindex = 0;
                    foreach (ListViewItem a in watListView1.Items)
                    {
                        if (a.SubItems[0].Text.Contains("AUX_BOILER"))
                        {
                            index = a.Index;
                            ME = a;
                        }
                        if (a.SubItems[0].Text.Contains("1E51") || a.SubItems[0].Text.Contains("2E51"))
                        {
                            moveindex = a.Index;
                        }
                    }
                    oItem = (ListViewItem)watListView1.Items[index].Clone();

                    if (index != 0 && moveindex != 0 && ME != null)
                    {
                        MoveItem(index, moveindex, oItem);
                    }
                    //한사이클 MA210
                    ME = null;
                    index = 0; moveindex = 0;
                    foreach (ListViewItem a in watListView1.Items)
                    {
                        if (a.SubItems[0].Text.Contains("MA210"))
                        {
                            index = a.Index;
                            ME = a;
                        }
                        if (a.SubItems[0].Text.Contains("2E31")|| a.SubItems[0].Text.Contains("1E31"))//2E31"))
                        {
                            moveindex = a.Index;
                        }
                    }
                    oItem = (ListViewItem)watListView1.Items[index].Clone();

                    if (index != 0 && moveindex != 0 && ME != null)
                    {
                        try
                        {
                            MoveItem(index, moveindex - 1, oItem);
                        }
                        catch { }
                    }
                    //한사이클 MA230
                    ME = null;
                    index = 0; moveindex = 0;
                    foreach (ListViewItem a in watListView1.Items)
                    {
                        if (a.SubItems[0].Text.Contains("MA230"))
                        {
                            index = a.Index;
                            ME = a;
                        }
                        if (a.SubItems[0].Text.Contains("2E41")|| a.SubItems[0].Text.Contains("1E41"))
                        {
                            moveindex = a.Index;
                        }
                    }
                    oItem = (ListViewItem)watListView1.Items[index].Clone();

                    if (index != 0 && moveindex != 0 && ME != null)
                    {
                        try
                        {
                            MoveItem(index, moveindex - 1, oItem);
                        }
                        catch { }
                    }
                    //한사이클 MA240
                    ME = null;
                    index = 0; moveindex = 0;
                    foreach (ListViewItem a in watListView1.Items)
                    {
                        if (a.SubItems[0].Text.Contains("MA240"))
                        {
                            index = a.Index;
                            ME = a;
                        }
                        if (a.SubItems[0].Text.Contains("2E51")|| a.SubItems[0].Text.Contains("1E51"))
                        {
                            moveindex = a.Index;
                        }
                    }
                    oItem = (ListViewItem)watListView1.Items[index].Clone();

                    if (index != 0 && moveindex != 0 && ME != null)
                    {
                        try
                        {
                            MoveItem(index, moveindex - 1, oItem);
                        }
                        catch { }
                    }
                    //한사이클 MA280
                    ME = null;
                    index = 0; moveindex = 0;
                    bool isfirst = true;
                    foreach (ListViewItem a in watListView1.Items)
                    {
                        if (a.SubItems[0].Text.Contains("MA280"))
                        {
                            index = a.Index;
                            ME = a;
                        }
                        if (a.SubItems[0].Text.Contains("1H11") || a.SubItems[0].Text.Contains("1H12") && isfirst)
                        {
                            moveindex = a.Index;
                            isfirst = false;
                        }
                    }
                    oItem = (ListViewItem)watListView1.Items[index].Clone();

                    if (index != 0 && moveindex != 0 && ME != null)
                    {
                        try
                        {
                            MoveItem(index, moveindex - 1, oItem);
                        }
                        catch { }
                    }
                }
            }
            catch { }
            //watListView1.Columns[2].TextAlign = HorizontalAlignment.Center;
            watListView1.Refresh();

            //SetHeight(watListView1, 20); //화면에 출력되는 watListView1 행 높이에 맞게 조절

            Console.WriteLine("End, ListViewUpdates");
        }
        private void MoveItem(int index,int moveindex, ListViewItem oItem)
        {
            watListView1.BeginUpdate();

            // 제거
            watListView1.Items[index].Remove();

            // 추가
            watListView1.Items.Insert(moveindex+1,oItem);

            watListView1.EndUpdate();
        }


        private List<string> ReadTxt(string dir)
        {
            string[] lines;
            using (var reader = File.OpenText(dir))
                lines = File.ReadAllLines(dir, Encoding.Default);

            List<string> treelist = new List<string>();

            foreach (string txt in lines)
            {
                if (lines[0].Contains('/'))
                {
                    string[] tt = txt.Split('/');
                    if (tt.Length == 3)
                        treelist.Add(tt[2]);
                }
                else if (lines[0].Contains(','))
                    treelist.Add(txt);
                else
                    treelist.Add(txt);
            }
            return treelist;
        }

        private void watListView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //
            string selectedItems = "";
            string selectedItem = "";
            bool addcomma = false;
            for (int i = 0; i < watListView1.SelectedItems.Count; i++)
            {
                selectedItem = watListView1.SelectedItems[i].SubItems[0].Text;

                if (selectedItem == "") { continue; }

                if (addcomma)
                {
                    selectedItems += "," + selectedItem;
                }
                else
                {
                    selectedItems += selectedItem;
                    addcomma = true;
                }
                if((i+1)%6==0)
                {
                    selectedItems += Environment.NewLine;
                    addcomma = false;
                }
            }
            if (selectedItems.Length > 1) selectedItems.Substring(1, selectedItems.Length - 1);
            label5.Text = String.Format("{0}", selectedItems);
            //
        }

        private void watListView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            if (this.watListView1.Sorting == SortOrder.Ascending || watListView1.Sorting == SortOrder.None)
            {
                this.watListView1.ListViewItemSorter = new ListViewItemComparer(e.Column, "desc");
                watListView1.Sorting = SortOrder.Descending;
            }
            else
            {
                this.watListView1.ListViewItemSorter = new ListViewItemComparer(e.Column, "asc");
                watListView1.Sorting = SortOrder.Ascending;
            }
            watListView1.Sort();
        }

        private void buttonUp_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Start, buttonUp_Click");

            if (watListView1.SelectedItems.Count != 1) { return; }

            for (int i = 0; i < watListView1.Items.Count; i++)
            {
                if (i == 0)
                {
                    if (watListView1.Items[i].SubItems[0].Text == watListView1.SelectedItems[0].SubItems[0].Text)
                    {
                        watListView1.Focus(); //
                        watListView1.Items[i].Selected = true;

                        return;
                    }
                }
            }

            bool bUp = true;
            listViewUpDown(bUp);

            Console.WriteLine("End, buttonUp_Click");
        }

        private void buttonDown_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Start, buttonDown_Click");

            if (watListView1.SelectedItems.Count != 1) { return; }

            for (int i = 0; i < watListView1.Items.Count; i++)
            {
                if (i == watListView1.Items.Count - 1)
                {
                    if (watListView1.Items[i].SubItems[0].Text == watListView1.SelectedItems[0].SubItems[0].Text)
                    {
                        watListView1.Focus(); //
                        watListView1.Items[i].Selected = true;

                        return;
                    }
                }
            }
            bool bUp = false;
            listViewUpDown(bUp);

            Console.WriteLine("End, buttonDown_Click");
        }

        private void listViewUpDown(bool bUp)
        {
            Console.WriteLine("Start, listViewUpDown");

            int iAfterSelected = 0;

            //
            listViewTemp.Items.Clear();
            Console.WriteLine("watListView1.Items.Count:{0}", watListView1.Items.Count);
            for (int i = 0; i < watListView1.Items.Count; i++)
            {
                bool bSame = false;
                string sSort = String.Format("{0:D3}.5", i);

                int iNEW = i;

                //////Console.WriteLine("-watListView1.Items[{0}].SubItems[0].Text:{1}", i, watListView1.Items[i].SubItems[0].Text);
                if (watListView1.Items[i].SubItems[0].Text == watListView1.SelectedItems[0].SubItems[0].Text)
                {
                    //////Console.WriteLine(" =watListView1.Items[{0}].SubItems[0].Text:{1}", i, watListView1.Items[i].SubItems[0].Text);
                    bSame = true;
                    if (bUp == true)
                    {
                        sSort = String.Format("{0:D3}.3", i - 1);
                        iNEW = int.Parse(String.Format("{0}", i - 1));

                    }
                    else
                    {
                        sSort = String.Format("{0:D3}.7", i + 1);
                        iNEW = int.Parse(String.Format("{0}", i + 1));
                    }
                    iAfterSelected = int.Parse(String.Format("{0}", iNEW));
                    Console.WriteLine("iAfterSelected:{0}", iAfterSelected);
                }

                string[] strArrs = new string[maxCols];
                for (int j = 0; j < maxCols; j++)
                {
                    //Console.WriteLine("j:{0}", j);
                    if (j == 0) { strArrs[j] = watListView1.Items[i].SubItems[j].Text; }
                    else if (j == 1) { strArrs[j] = watListView1.Items[i].SubItems[j].Text; }
                    else if (j == 2) { strArrs[j] = watListView1.Items[i].SubItems[j].Text; }
                    else if (j == 3) { strArrs[j] = String.Format("{0:D4}", i); }
                    else if (j == 4) { strArrs[j] = watListView1.Items[i].SubItems[j].Text; }
                    else if (j == 5) { strArrs[j] = watListView1.Items[i].SubItems[j].Text; }
                    else if (j == 6) { strArrs[j] = String.Format("{0}", sSort); }
                    else if (j == 7) { strArrs[j] = watListView1.Items[i].SubItems[j].Text; ; }
                    else { strArrs[j] = ""; }
                }
                ListViewItem strArr = new ListViewItem(strArrs);
                listViewTemp.Items.Add(strArr);
            }
            //
            this.listViewTemp.ListViewItemSorter = new ListViewItemComparer(5, "asc");
            listViewTemp.Sorting = SortOrder.Ascending;
            listViewTemp.Sort();
            listViewTemp.Refresh();

            /////////////////////////////////////////////

            Console.WriteLine("listViewTemp.Items.Count:{0}", listViewTemp.Items.Count);
            watListView1.Items.Clear();
            for (int i = 0; i < listViewTemp.Items.Count; i++)
            {
                //////Console.WriteLine("-listViewTemp.Items[{0}].SubItems[0].Text:{1}", i, listViewTemp.Items[i].SubItems[0].Text);

                string[] strArrs = new string[maxCols];
                for (int j = 0; j < maxCols; j++)
                {
                    //Console.WriteLine("j:{0}", j);
                    strArrs[j] = listViewTemp.Items[i].SubItems[j].Text;
                }
                ListViewItem strArr = new ListViewItem(strArrs);
                watListView1.Items.Add(strArr);
            }
            Console.WriteLine("===");
            //
            this.watListView1.ListViewItemSorter = new ListViewItemComparer(6, "asc");
            watListView1.Sorting = SortOrder.Ascending;
            watListView1.Sort();
            watListView1.Focus(); //
            watListView1.Refresh();

            watListView1.Items[iAfterSelected].Selected = true;
            watListView1.Items[iAfterSelected].Focused = true;
            watListView1.Focus();
            watListView1.Items[iAfterSelected].EnsureVisible();

            Console.WriteLine("End, listViewUpDown");
        }

        private void buttonSaveErectNo_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Start, buttonSaveErectNo_Click. _erectSavedFile:{0}", _erectSavedFile);
            if(allblock)
            {
                _erectSavedFile = Path.Combine(_erectPath, @"Z98_ERECT_ALLBLOCK_SAVED.TXT");
            }
            FileStream fileStreamOutputATTTemp = new FileStream(_erectSavedFile, FileMode.Create);
            fileStreamOutputATTTemp.Seek(0, SeekOrigin.Begin);

            _savedUser = String.Format("{0}", _userName);

            string sI = "";
            byte[] info;

            sI = String.Format("#USERNAME,{0}", _savedUser);
            info = System.Text.Encoding.Default.GetBytes(sI + "\r\n"); //ANSI
            //////info = new UTF8Encoding(true).GetBytes(sI + "\r\n"); //UTF-8

            fileStreamOutputATTTemp.Write(info, 0, info.Length);

            for (int i = 0; i < watListView1.Items.Count; i++)
            {
                string erectBLK = watListView1.Items[i].SubItems[0].Text;
                string erectNo = String.Format("{0:D3}.0", i);
                string erectDir = watListView1.Items[i].SubItems[7].Text;
                string erectLength = watListView1.Items[i].SubItems[8].Text;

                sI = String.Format("{0},{1},{2},{3}", erectBLK, erectNo, erectDir, erectLength);
                info = System.Text.Encoding.Default.GetBytes(sI + "\r\n"); //ANSI
                //////info = new UTF8Encoding(true).GetBytes(sI + "\r\n"); //UTF-8

                fileStreamOutputATTTemp.Write(info, 0, info.Length);
            }
            fileStreamOutputATTTemp.Write(info, 0, info.Length);
            fileStreamOutputATTTemp.Flush();
            fileStreamOutputATTTemp.Close();

            label11.Text = String.Format("탑재일자 수정 ({0})", _savedUser);

            MessageBox.Show(String.Format("[Succeeded]Saved File : {0}", _erectSavedFile));

            cbBox1.Text = cbBox1.SelectedItem.ToString();

            Console.WriteLine("End, buttonSaveErectNo_Click");
        }

        private void buttonScreenCSV_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Start, buttonScreenCSV_Click");

            //_excelFile
            int TitleLines = 1;

            string excelFileSrc = Path.Combine(_originPath, String.Format("screen.xlsx"));
            string targetFile = Path.Combine(_tempDir, String.Format("{0}_screen.xlsx", _shipNo));
            if (!File.Exists(excelFileSrc)) // Excel 화일(복사 대상)
            {
                MessageBox.Show(String.Format("[Error]Not Exist Excel Origin File({0})", excelFileSrc));
                return;
            }
            try { File.Copy(excelFileSrc, targetFile, true); }
            catch
            {
                MessageBox.Show(String.Format("[Error]Copy Excel Origin File({0}) -> targetFile:({1})", excelFileSrc, targetFile));
                return;
            }

            Console.WriteLine("@@@@@@@@@@@@@@@@@@@@@@@@@@@@");

            object TypMissing = Type.Missing;
            Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            string excelVersion = ExcelApp.Version.Substring(0, 2).ToString();
            //MessageBox.Show(String.Format("Excel : {0} >>> [14->2010,12->2007,11->2003...]", excelVersion));
            Excel.Workbook _workbook = null;

            _workbook = ExcelApp.Workbooks.Open(targetFile, TypMissing, TypMissing, TypMissing, TypMissing,
                    TypMissing, TypMissing, TypMissing, TypMissing, TypMissing, TypMissing, TypMissing, TypMissing, TypMissing, TypMissing);

            Excel.Worksheet Sheet = (Excel.Worksheet)_workbook.Worksheets.get_Item("1"); //Sheet1
            Excel.Range Range_ = Sheet.get_Range("A1", Type.Missing);

            try
            {
                for (int i = 0; i < watListView1.Items.Count; i++)
                {
                    for (int j = 0; j < watListView1.Items[i].SubItems.Count; j++)
                    {
                        //if (j >= 7)
                        //    continue;
                        
                        if(j>1)
                        {
                            Sheet.Cells[TitleLines + i + 1, j ] = watListView1.Items[i].SubItems[j].Text.Trim();
                        }
                        else
                        {
                            Sheet.Cells[TitleLines + i + 1, j + 1] = watListView1.Items[i].SubItems[j].Text.Trim();
                        }
                    }
                }

                Microsoft.Office.Interop.Excel.Range range = null;
                range = Sheet.get_Range(String.Format("A1"), String.Format("F{0}", watListView1.Items.Count + 1));
                range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                range.Borders.Weight = Excel.XlBorderWeight.xlThin;
                //range.BorderAround(TypMissing, Excel.XlBorderWeight.xlThick, Excel.XlColorIndex.xlColorIndexAutomatic, TypMissing);

                //Sheet.Columns.AutoFit();

            }
            catch
            {
            }
            finally
            {
                _workbook.Save();
                ExcelApp.Workbooks.Close();
                ExcelApp.Quit();
            }

            buttonTempDir_Click(null, null);

            Console.WriteLine("End, buttonScreenCSV_Click");
        }

        private void buttonTempDir_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Start, buttonTempDir_Click");

            string exeFile = String.Format(@"explorer.exe ""{0}""", _tempDir);
            RunCMDOnly(exeFile, true, false);

            Console.WriteLine("End, buttonTempDir_Click");
        }

        private void buttonShipDATA_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Start, buttonShipDATA_Click");

            string exeFile = String.Format(@"explorer.exe ""{0}""", _selectedSHIPPath);
            RunCMDOnly(exeFile, true, false);

            Console.WriteLine("End, buttonShipDATA_Click");
        }
        private List<ClashData> Clash = new List<ClashData>();
        Stopwatch lugtime = new Stopwatch();
        private void button2_Click(object sender, EventArgs e)
        {
            Connector.OnFinishedClashTestEvent += Connector_LUG_FinishedClashTestEvent;
            timer_RUG.Enabled = false;
            timer_RUG.Tick += timer_RUG_Tick;

            lugtime.Restart();

            Console.WriteLine("Start, button2_Click");
            foreach (ListViewItem a in listView2.SelectedItems)
            {
                ClashData data = new ClashData();
                data.BlockName = a.SubItems[0].Text;
                data.BlockPath = a.SubItems[6].Text;
                Clash.Add(data);
            }

            LugClashRun();
            Console.WriteLine("End, button2_Click");
        }
        List<string> tree = new List<string>();
        public void LugClashRun()
        {
            ClashData clash = null;
            for (int i = 0; i < Clash.Count; i++)
            {
                ClashData vo = Clash[i];
                if (vo.Status == ClashStatus.TESTED)
                {
                    try
                    {
                        Connector.DeleteClashTest(vo.clash.taskID, vo.clash.translationID);
                    }
                    catch (Exception ex) { MessageBox.Show(ex.Message); }
                }
                if (vo.Status == ClashStatus.NONE)
                {
                    clash = vo;
                    break;
                }
            }

            if (clash == null)
            {
                Connector.OnFinishedClashTestEvent -= Connector_LUG_FinishedClashTestEvent;
                string time = lugtime.Elapsed.Hours + "시" + lugtime.Elapsed.Minutes + "분" + lugtime.Elapsed.Seconds + "초";
                if (autorun)
                {
                    Connector.Exit(true);
                }
                MessageBox.Show(string.Format("Lug&Wire 간섭체크가 끝났습니다. 총 소요시간:{0}", time));
                return;
            }

            clash.Status = ClashStatus.TESTING;
            Connector.ShowWaitDialogWithText(true, "LUG & WIRE 간섭체크", string.Format("{0}, {1}/{2}", clash.BlockName,Clash.IndexOf(clash)+1, Clash.Count()));

            Connector.IgnoreModelChangedStatus(true);
            Connector.CloseDocument();
            List<string> paths = new List<string>();
            string currentblock = clash.BlockName;
            string blockPath = Path.Combine(_tempDir2, currentblock + ".VIZXML");
            string block = "";
            bool isfirst = true;
            foreach (var a in tree)
            {
                if (a.Contains("/" + currentblock + "/"))
                {
                    if (isfirst)
                    {
                        block = a.Split('/').ToList()[2];
                        isfirst = false;
                    }
                }
            }
            //////////////
            StreamWriter writer1 = new StreamWriter(Path.Combine(_tempDir2, currentblock+".VIZXML"), false);
            string path = _originPath; string filepath = "";

            bool first = true;

            foreach (var a in tree)
            {
                if (a.Split('/').Last() == currentblock && first)
                {
                    if (a.Split('/')[a.Split('/').Count() - 2].Length == 4 || a.Split('/')[a.Split('/').Count() - 2] == cbBox1.Text)
                    {
                        filepath = a;
                        first = false;
                    }
                }
            }
            filepath = filepath.Substring(cbBox1.Text.Length + 1);
            //filepath = filepath.Substring(block.Length + 1); 하위에 블록이 없을시 제거기능
            filepath = filepath.Replace('/', '\\');

            writer1.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            writer1.WriteLine("<VIZXML>");
            writer1.WriteLine("<Model Name=\"{0}\" SkipBrokenLinks=\"True\" UncheckToUnload=\"False\">", cbBox1.Text);
            string ifvizf = Path.Combine(path, cbBox1.Text, block, block + "_정적.viz");
            if (block == currentblock)
            {
                if (new FileInfo(ifvizf).Exists)
                {
                    writer1.WriteLine("<Node Name=\"{2}\" ExtLinkFile=\"{3}\\{0}\\{1}\\{1}_정적.viz\" HideAndLock=\"False\" UncheckToUnload=\"False\" Type=\"Assembly\"/>",
                        cbBox1.Text, block, currentblock, path);
                }
                else
                {
                    writer1.WriteLine("<Node Name=\"{2}\" ExtLinkFile=\"{3}\\{0}\\{1}\\{1}.rev\" HideAndLock=\"False\" UncheckToUnload=\"False\" Type=\"Assembly\"/>",
    cbBox1.Text, block, currentblock, path);
                }
            }
            else//re
            {
                if (new FileInfo(ifvizf).Exists)
                {
                    writer1.WriteLine("<Node Name=\"{3}\" ExtLinkNode=\"{4}\\{0}\\{1}\\{1}_정적.viz:{2}\" HideAndLock=\"False\" UncheckToUnload=\"False\" Type=\"Assembly\"/>",
                    cbBox1.Text, block, filepath, currentblock, path);
                }
                else
                {
                    writer1.WriteLine("<Node Name=\"{3}\" ExtLinkNode=\"{4}\\{0}\\{1}\\{1}.rev:{2}\" HideAndLock=\"False\" UncheckToUnload=\"False\" Type=\"Assembly\"/>",
cbBox1.Text, block, filepath, currentblock, path);
                }
            }

            writer1.WriteLine("</Model>");
            writer1.WriteLine("</VIZXML>");
            writer1.Close();
            //////////////

            string lugPath = Path.Combine(_originPath, _shipNo, "LUG_WIRE", "LUG_WIRE.viz");
            if(!new FileInfo(lugPath).Exists)
            {
                lugPath = Path.Combine(_originPath, _shipNo, "LUG_WIRE", "LUG_WIRE.REV");
            }
            paths.Add(blockPath); paths.Add(lugPath);
            Connector.AddDocuments(paths.ToArray());
            List<int> blockID = new List<int>();
            List<int> facID = new List<int>();
            List<int> shownode = new List<int>();
            //
            foreach (var b in Connector.GetAllObjects())
            {
                if (b.NodeName == currentblock)
                {
                    blockID.Add(b.Id);
                    shownode.Add(b.Index);
                }
                if (b.NodeName == "LUG_WIRE_" + currentblock)
                {
                    facID.Add(b.Id);
                    shownode.Add(b.Index);
                }
            }
            

            Connector.ShowObjects(shownode.ToArray(), true);
            clash.clash = Connector.AddClashTest();
            setResultType(); setResultOption();

            Connector.EnableMessageBox(MESSAGEBOX_TYPES.CLASHTEST_COMPLETED, false);
            Connector.StartClashCheck(blockID.ToArray(), facID.ToArray());
        }
        private void Connector_LUG_FinishedClashTestEvent(object sender, FinishedClashTestEventArgs e)
        {
            Connector.ShowWaitDialog(false);
            ClashData clash = null;
            for (int i = 0; i < Clash.Count; i++)
            {
                ClashData vo = Clash[i];
                if (vo.clash != null)
                {
                    if (vo.clash.taskID == e.TaskID)
                        clash = vo;
                }
            }
            clash.Status = ClashStatus.TESTED;
            //간섭결과 가공


            List<ClashResultVO> ClashResult = Connector.GetClashTestResultList(0, 0);
            int count = 1;

            //상위 블록 가져오기
            string block = "";
            bool isfirst = true;
            foreach (var a in tree)
            {
                if (a.Contains("/" + clash.BlockName + "/"))
                {
                    if (isfirst)
                    {
                        block = a.Split('/').ToList()[2];
                        isfirst = false;
                    }
                }
            }
            if (block == "")
            {
                timer_RUG.Enabled = true;
                timer_RUG.Start();
                MessageBox.Show(string.Format("{0}의 상위블록을 찾지 못했습니다.", clash.BlockName));
                return;
            }
            string path = Path.Combine(dic_treenode2Path[block], clash.BlockName + "_간섭체크11.CSV");
            StreamWriter SW = new StreamWriter(path, false,Encoding.UTF8);
            foreach (var b in ClashResult)
            {
                ClashResultVO result = b;
                string write =(String.Format("{0}∫{1}∫{2}∫{3}∫{4}∫{5}∫{6}∫{7}∫{8}∫{9}∫{10}∫{11}∫{12}∫{13}∫{14}∫{15}∫{16}∫{17}∫{18}∫{19}∫{20}∫{21}∫{22}∫{23}∫{24}",
                    count, //0
                    clash.BlockName, //1
                    "LUG_WIRE", //2
                    result.PartNodeA, //3 (Model1)
                    result.PartNodeB, //4 (Modle2)
                    result.ResultType, //5
                    result.Distance, //6
                    result.PointX, //7
                    result.PointY, //8
                    result.PointZ, //9
                    "", //panelName1, //10
                    "", //panelName2, //11
                    "CtTypeG2G", //12
                    result.ClashResultString, //13
                    "", //modelDept1, //14
                    "", //modelDept2, //15
                    "", //modelType1, //16
                    "", //modelType2, //17
                    "", //18 O X
                    "", //sosokBLK, //19
                    result.NodePathA, //20
                    result.NodePathB, //21
                    "", //moduleORassy1, //22
                    "", //moduleORassy2, //23
                    "" //24
                    ));
                count++;
                SW.WriteLine(write);
            }
            SW.Close();

            timer_RUG.Enabled = true;
            timer_RUG.Start();
        }
        private void timer_RUG_Tick(object sender, EventArgs e)
        {
            timer_RUG.Enabled = false;
            timer_RUG.Stop();

            Connector.ShowWaitDialog(false);
            LugClashRun();           
        }
        public void setResultType()
        {
            ClashResultOption vo = Connector.GetClashResultOption();

            vo.ViewAssembly = false;
            vo.SurroundViewMode = 1;// 1-주변보기 2-그것만 보기
            vo.RangePercent = 3;
            vo.SurroundRenderType = 2;
            vo.UseSingleColor = true;
            vo.SingleColor = Color.DeepSkyBlue;
            vo.UseChangeColor = true;
            vo.Item1Color = Color.WhiteSmoke;
            vo.Item2Color = Color.Red;
            vo.ResultRenderType = 2;
            vo.ShowHotPoint = false;
            vo.ResultViewType = 2;
            vo.ResultViewRate = 2;
            vo.PointViewType = 0; //간섭결과 보기 0 전체, 1 고정, 2 확대

            Connector.SetClashResultOption(vo);
        }
        public void setResultOption()
        {
            ClashOption vo = Connector.GetClashOption();
            vo.ClashType = CtType.CtTypeG2G;
            
            vo.VisibleOnly = true;
            vo.UseToleranceRange = true;//근접허용오차
            vo.UseCalibrationTolerance = true;//접촉허용오차
            vo.ToleranceRange = float.Parse(textBox1.Text);
            vo.CalibrationTolerance = float.Parse(textBox2.Text);
            vo.ExceptLevel = int.Parse(textBox3.Text);
            Connector.SetClashOption(vo);
        }

        public void AUTOClashRun()
        {
            try
            {
                ClashData clash = null;
                for (int i = 0; i < Clash.Count; i++)
                {
                    ClashData vo = Clash[i];
                    if (vo.Status == ClashStatus.TESTED)
                    {
                        try
                        {
                            //Connector.DeleteClashTest(vo.clash.taskID, vo.clash.translationID);
                        }
                        catch (Exception ex) { MessageBox.Show(ex.Message); }
                    }
                    if (vo.Status == ClashStatus.NONE)
                    {
                        clash = vo;
                        break;
                    }
                }

                if (clash == null)
                {
                    Connector.OnFinishedClashTestEvent -= Connector_AUTO_FinishedClashTestEvent;
                    string time = lugtime.Elapsed.Hours + "시" + lugtime.Elapsed.Minutes + "분" + lugtime.Elapsed.Seconds + "초";
                    if (autorun)
                    {
                        Connector.Exit(true);
                    }
                    return;
                }

                clash.Status = ClashStatus.TESTING;
                Connector.ShowWaitDialogWithText(true, "LUG & WIRE 간섭체크", string.Format("{0}, {1}/{2}", clash.BlockName, Clash.IndexOf(clash) + 1, Clash.Count()));

                //Connector.IgnoreModelChangedStatus(true);
                //Connector.CloseDocument();
                List<string> paths = new List<string>();
                string currentblock = clash.BlockName;
                List<string> add = new List<string>() { currentblock };
                Connector.AddDocuments(add.ToArray());
                List<int> blockID = new List<int>();
                //
                foreach (var b in Connector.GetAllObjects())
                {
                    if (b.NodeName == currentblock.Split('\\').Last())
                    {
                        blockID.Add(b.Id);
                    }
                }
                clash.clash = Connector.AddClashTest();
                setResultType2(); setResultOption2();

                Connector.EnableMessageBox(MESSAGEBOX_TYPES.CLASHTEST_COMPLETED, false);
                Connector.ShowAll(true);
                Connector.StartClashCheck(blockID.ToArray(), blockID.ToArray());
            }
            catch(Exception ea) { MessageBox.Show(ea.Message); }
        }
        private void Connector_AUTO_FinishedClashTestEvent(object sender, FinishedClashTestEventArgs e)
        {
            Connector.ShowWaitDialog(false);
            ClashData clash = null;
            for (int i = 0; i < Clash.Count; i++)
            {
                ClashData vo = Clash[i];
                if (vo.clash != null)
                {
                    if (vo.clash.taskID == e.TaskID)
                        clash = vo;
                }
            }
            clash.Status = ClashStatus.TESTED;
            ////간섭결과 가공


            //List<ClashResultVO> ClashResult = Connector.GetClashTestResultList(0, 0);
            //int count = 1;

            ////상위 블록 가져오기
            //string block = "";
            //bool isfirst = true;
            //foreach (var a in tree)
            //{
            //    if (a.Contains("/" + clash.BlockName + "/"))
            //    {
            //        if (isfirst)
            //        {
            //            block = a.Split('/').ToList()[2];
            //            isfirst = false;
            //        }
            //    }
            //}
            //if (block == "")
            //{
            //    timer_AUTO.Enabled = true;
            //    timer_AUTO.Start();
            //    MessageBox.Show(string.Format("{0}의 상위블록을 찾지 못했습니다.", clash.BlockName));
            //    return;
            //}
            //string path = Path.Combine(@"C:\temp" + "_간섭체크11.CSV");
            //StreamWriter SW = new StreamWriter(path, false, Encoding.UTF8);
            //foreach (var b in ClashResult)
            //{
            //    ClashResultVO result = b;
            //    string write = (String.Format("{0}∫{1}∫{2}∫{3}∫{4}∫{5}∫{6}∫{7}∫{8}∫{9}∫{10}∫{11}∫{12}∫{13}∫{14}∫{15}∫{16}∫{17}∫{18}∫{19}∫{20}∫{21}∫{22}∫{23}∫{24}",
            //        count, //0
            //        clash.BlockName, //1
            //        "LUG_WIRE", //2
            //        result.PartNodeA, //3 (Model1)
            //        result.PartNodeB, //4 (Modle2)
            //        result.ResultType, //5
            //        result.Distance, //6
            //        result.PointX, //7
            //        result.PointY, //8
            //        result.PointZ, //9
            //        "", //panelName1, //10
            //        "", //panelName2, //11
            //        "CtTypeG2G", //12
            //        result.ClashResultString, //13
            //        "", //modelDept1, //14
            //        "", //modelDept2, //15
            //        "", //modelType1, //16
            //        "", //modelType2, //17
            //        "", //18 O X
            //        "", //sosokBLK, //19
            //        result.NodePathA, //20
            //        result.NodePathB, //21
            //        "", //moduleORassy1, //22
            //        "", //moduleORassy2, //23
            //        "" //24
            //        ));
            //    count++;
            //    SW.WriteLine(write);
            //}
            //SW.Close();

            timer_AUTO.Enabled = true;
            timer_AUTO.Start();
        }
        private void timer_AUTO_Tick(object sender, EventArgs e)
        {
            timer_AUTO.Enabled = false;
            timer_AUTO.Stop();

            Connector.ShowWaitDialog(false);
            AUTOClashRun();
        }
        public void setResultType2()
        {
            ClashResultOption vo = Connector.GetClashResultOption();

            vo.ViewAssembly = false;
            vo.SurroundViewMode = 0;// 1-주변보기 2-그것만 보기
            vo.RangePercent = 3;
            vo.SurroundRenderType = 2;
            vo.UseSingleColor = true;
            vo.SingleColor = Color.DeepSkyBlue;
            vo.UseChangeColor = true;
            vo.Item1Color = Color.Yellow;
            vo.Item2Color = Color.Red;
            vo.ResultRenderType = 2;
            vo.ShowHotPoint = false;
            vo.ResultViewType = 2;
            vo.ResultViewRate = 2;
            vo.PointViewType = 2; //간섭결과 보기 0 전체, 1 고정, 2 확대

            Connector.SetClashResultOption(vo);
        }
        public void setResultOption2()
        {
            ClashOption vo = Connector.GetClashOption();
            vo.ClashType = CtType.CtTypeSelf;

            vo.VisibleOnly = true;
            vo.UseToleranceRange = true;//근접허용오차
            vo.UseCalibrationTolerance = true;//접촉허용오차
            vo.ToleranceRange = float.Parse(textBox1.Text);
            vo.CalibrationTolerance = float.Parse(textBox2.Text);
            vo.ExceptLevel = int.Parse(textBox3.Text);
            Connector.SetClashOption(vo);
        }
        private void ClashControl_Load(object sender, EventArgs e)
        {
            Connector.SetShowAutoFitMode(false);
            ChangeBOX.Hide();
            if (companyName != "IFG")
                button4.Visible = false;
            this.watListView1.ComboChanged += new WATListView.eventComboChanged(lvwColor_ComboChanged);
            List<string> strItem = new List<string>();
            strItem.Add("X+");
            strItem.Add("X-");
            strItem.Add("Y+");
            strItem.Add("Y-");
            strItem.Add("Z");
            this.watListView1.AddString(7, strItem.ToArray());

            int VMajor = 3; int VMinor = 0; int VBuild = 2; int VRevision = 19072;
            string needver = VMajor.ToString() + "." + VMinor.ToString()
                + "." + VBuild.ToString() + "." + VRevision.ToString();
            try
            {
                bool verRight = true;
                string ver = Connector.GetVersion().Major + "." + Connector.GetVersion().Minor
                    + "." + Connector.GetVersion().Build + "." + Connector.GetVersion().Revision;

                if (ver.Length==11)
                {
                    if (Connector.GetVersion().Major > VMajor) { }
                    else if(Connector.GetVersion().Major < VMajor) { verRight = false; }
                    else
                    {
                        if (Connector.GetVersion().Minor > VMinor) { }
                        else if (Connector.GetVersion().Minor < VMinor) { verRight = false; }
                        else
                        {
                            if (Connector.GetVersion().Build > VBuild) { }
                            else if (Connector.GetVersion().Build < VBuild) { verRight = false; }
                            else
                            {
                                if (Connector.GetVersion().Revision >= VRevision) { }
                                else{ verRight = false; }
                            }
                        }
                    }
                        
                }
                if (!verRight)
                {
                    MessageBox.Show(string.Format("{0}이상의 버전의 VIZZARD가 필요합니다. 현재 버전:{1}", needver, ver)
                        ,"VIZZARD 버전이 낮습니다");
                }                   
            }
            catch
            {
                MessageBox.Show(string.Format("현재 VIZZARD 버전이 낮습니다. 필요 버전:{0}",needver));
            }

            try
            {
                if (companyName == "HMD")
                {
                    initfile();
                    DirectoryInfo local = new DirectoryInfo(LocalPath);
                    DirectoryInfo server = new DirectoryInfo(string.Format(_vizzardDir, "Plugins"));

                    List <FileInfo> serverList = server.GetFiles().ToList();
                    List<FileInfo> localList = local.GetFiles().ToList();

                    foreach (var a in serverList)
                    {
                        if (a.Name != "autoCLASH.dll")
                            continue;
                        bool checkfile = false;
                        foreach (var b in localList)
                        {
                            if (a.Name == b.Name)
                            {
                                checkfile = true;
                                if (a.LastWriteTime>b.LastWriteTime)
                                {
                                    MessageBox.Show(string.Format("server파일:{0}/수정날짜:{1} ↔ local파일:{2}/수정날짜:{3}의 버전이 낮습니다." +
                                        " 관리자에게 문의하세요."
                                        , a.FullName, a.LastWriteTime, b.FullName, b.LastWriteTime), "autoClash.dll 업데이트 필요");
                                    return;
                                }
                            }
                        }
                        //if (!checkfile)
                        //{
                        //    MessageBox.Show(string.Format("{0}의 파일이 Local에 없습니다. 관리자에게 문의하세요.", a.Name),"DLL 파일이 없음");
                        //    return;
                        //}
                    }
                }
                else
                {
                    //DirectoryInfo server = new DirectoryInfo(@"F:\test\server");//@"C:\VIZZARD64_Design");
                    //DirectoryInfo local = new DirectoryInfo(@"F:\test\local");//@"H:\HMD\C\update\VIZZARD64_Design");

                    //List<FileInfo> serverList = server.GetFiles().ToList();
                    //List<FileInfo> localList = local.GetFiles().ToList();

                    //foreach (var a in serverList)
                    //{
                    //    if (a.Name != "autoCLASH.dll")
                    //        continue;
                    //    bool checkfile = false;
                    //    foreach (var b in localList)
                    //    {
                    //        if (a.Name == b.Name)
                    //        {
                    //            checkfile = true;
                    //            if (a.LastWriteTime > b.LastWriteTime)
                    //            {
                    //                MessageBox.Show(string.Format("server파일:{0}/수정날짜:{1} ↔ local파일:{2}/수정날짜:{3}의 버전이 낮습니다." +
                    //                    " 관리자에게 문의하세요."
                    //                    , a.FullName, a.LastWriteTime, b.FullName, b.LastWriteTime), "autoClash.dll 업데이트 필요");
                    //                return;
                    //            }
                    //        }
                    //    }
                    //    //if (!checkfile)
                    //    //{
                    //    //    MessageBox.Show(string.Format("{0}의 파일이 Local에 없습니다. 관리자에게 문의하세요.", a.Name),"DLL 파일이 없음");
                    //    //    return;
                    //    //}
                    //}
                }
            }
            catch { }

            DateTime checkTime = DateTime.Now;
            decimal timeTemp = decimal.Parse(String.Format("{0:D4}{1:D2}{2:D2}{3:D2}{4:D2}{5:D2}.{6}", checkTime.Year, checkTime.Month, checkTime.Day, checkTime.Hour, checkTime.Minute, checkTime.Second, checkTime.Millisecond));
            if (companyName == "IFG") { allowTimeMAX = 20501231235959.99m; }
            if (companyName == "HMD") { allowTimeMAX = 20991231235959.99m; }
            if (companyName == "HHISS") { allowTimeMAX = 20991231235959.99m; }

            System.Console.WriteLine("allowTimeMAX:{0}", allowTimeMAX);
            if (timeTemp < allowTimeMIN)
            {
                MessageBox.Show("License Expired...");
                Connector.Exit(true);
            }
            if (timeTemp > allowTimeMAX)
            {
                MessageBox.Show("License Expired...");
                Connector.Exit(true);
            }
            string licesneExpiredDate = String.Format("{0}-{1}-{2}", allowTimeMAX.ToString().Substring(0, 4), allowTimeMAX.ToString().Substring(4, 2), allowTimeMAX.ToString().Substring(6, 2));
            this.Text += string.Format("[License expired : {0}]", licesneExpiredDate);
            //this.Close();   
        }
        public static decimal allowTimeMIN = 20190101000000.00m;
        public static decimal allowTimeMAX = 20190202235959.99m;
        void lvwColor_ComboChanged(WATListComboEventArgs e)
        {
            //Trace.Write("lvwColor_ComboChanged : " + e.Combobox.Text);
            //switch (e.Combobox.Text)
            //{
            //    case "빨강":
            //        e.SelectedItem.ForeColor = Color.Red;
            //        break;
            //    case "노랑":
            //        e.SelectedItem.ForeColor = Color.Yellow;
            //        break;
            //    case "파랑":
            //        e.SelectedItem.ForeColor = Color.Blue;
            //        break;
            //    case "초록":
            //        e.SelectedItem.ForeColor = Color.Green;
            //        break;
            //}
        }
        public void initfile()
        {            
            StreamReader sr00 = new StreamReader(initpath);
            string line00 = "";
            while (sr00.EndOfStream == false)
            {
                line00 = sr00.ReadLine().Trim();
                if (line00.Length < 3) { continue; }
                if (line00.Substring(0, 1) == "[") { continue; }
                if (line00.Substring(0, 1) == "!") { continue; }
                if (line00.IndexOf("!") > 0)
                {
                    line00 = line00.Substring(0, line00.IndexOf("!")).Trim();
                }
                if (line00.IndexOf("VIZZARD_DIR=") == 0)
                {
                    string _s32_64 = "64";
                    string vizzardDirTemp = line00.Substring("VIZZARD_DIR=".Length);

                    string processorArchitecture = "";
                    try { processorArchitecture = Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE", EnvironmentVariableTarget.Machine); } //Machine 중요.
                    catch { }
                    if (companyName == "HMD")
                    {
                        _vizzardDir = Path.Combine(vizzardDirTemp, String.Format("VIZZARD{0}_Design", _s32_64));
                    }
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach(ListViewItem a in watListView1.Items)
            {
                a.SubItems[7].Text = comboBox1.Text;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }
        public bool preCheck()
        {
            string savepath = @"C:\temp\newtest.txt";
            StreamWriter sws = new StreamWriter(savepath);
            sws.WriteLine("테스트 진행");
            sws.Close();

            Stopwatch sw = new Stopwatch();
            sw.Start();
            //탑재방향 담아두기
            Dictionary<string, string> dir = new Dictionary<string, string>();
            foreach (ListViewItem a in watListView1.Items)
            {
                dir.Add(a.SubItems[0].Text, a.SubItems[7].Text);
            }

            List<NodeVO> items = Connector.GetChildObjects(blockindex, ChildrenTypes.Children);
            Connector.ShowWaitDialogWithText(true, "탑재방향 정합성 확인", "탑재방향 검사 중");
            Connector.ShowAll(false);
            int cc = 0;
            foreach (var a in items)
            {
                cc++;
                int cas = 0;
                Connector.ShowObject(a.Index, true);
                string blockname = a.NodeName.Replace(".rev", "");
                Connector.UpdateWaitDialogDescription(string.Format("{0}_탑재방향 검사 중 {1}/{2}", blockname, cc,items.Count));
                ObjectPropertyVO op = Connector.GetObjectProperty(a.Index, false);
                if (!dir.Keys.Contains(blockname))
                { continue; }
                Connector.SetSelected(0, false, false);
                Connector.SetSelected(a.Index, true, false);
                float Xlen = float.Parse(op.MaxPoint.X) - float.Parse(op.MinPoint.X);
                float Ylen = float.Parse(op.MaxPoint.Y) - float.Parse(op.MinPoint.Y);
                float Zlen = float.Parse(op.MaxPoint.Z) - float.Parse(op.MinPoint.Z);
                List<float> boxori = new List<float>();
                boxori.Add(Convert.ToSingle(op.MinPoint.X));
                boxori.Add(Convert.ToSingle(op.MinPoint.Y));
                boxori.Add(Convert.ToSingle(op.MinPoint.Z));
                boxori.Add(Convert.ToSingle(op.MaxPoint.X));
                boxori.Add(Convert.ToSingle(op.MaxPoint.Y));
                boxori.Add(Convert.ToSingle(op.MaxPoint.Z));

                float path = Zlen;
                int errorcount = 0;

                if (dir[blockname] == "X-")
                {
                    path = Xlen;
                    for (int i = 1; i < 5; i++)
                    {
                        List<float> box = new List<float>() { boxori[0]+Xlen*i/4, boxori[1], boxori[2], boxori[3] + Xlen * i / 4, boxori[4], boxori[5] };
                        List<int> searchItems = Connector.GetObjectsInArea(box.ToArray(), new int[] { }, CrossBoundBox.IncludingPart);//IncludingPart
                        List<NodeVO> test = Connector.GetChildObjects(a.Index, ChildrenTypes.All_Children);
                        List<int> test2 = new List<int>();
                        foreach (var abc in test)
                        {
                            test2.Add(abc.Index);
                        }
                        searchItems = searchItems.Except(test2).ToList();
                        cas = test2.Count;
                        errorcount += searchItems.Count();
                    }
                }
                else if (dir[blockname] == "X+")
                {
                    path = -Xlen;
                    for (int i = 1; i < 5; i++)
                    {
                        List<float> box = new List<float>() { boxori[0] - Xlen * i / 4, boxori[1], boxori[2], boxori[3] - Xlen * i / 4, boxori[4], boxori[5] };
                        List<int> searchItems = Connector.GetObjectsInArea(box.ToArray(), new int[] { }, CrossBoundBox.IncludingPart);//IncludingPart
                        List<NodeVO> test = Connector.GetChildObjects(a.Index, ChildrenTypes.All_Children);
                        List<int> test2 = new List<int>();
                        foreach (var abc in test)
                        {
                            test2.Add(abc.Index);
                        }
                        searchItems = searchItems.Except(test2).ToList();
                        cas = test2.Count;
                        errorcount += searchItems.Count();
                    }
                }
                else if (dir[blockname] == "Y-")
                {
                    path = Ylen;
                    for (int i = 1; i < 5; i++)
                    {
                        List<float> box = new List<float>() { boxori[0] , boxori[1] + Ylen * i / 4, boxori[2], boxori[3] , boxori[4] + Ylen * i / 4, boxori[5] };
                        List<int> searchItems = Connector.GetObjectsInArea(box.ToArray(), new int[] { }, CrossBoundBox.IncludingPart);//IncludingPart
                        List<NodeVO> test = Connector.GetChildObjects(a.Index, ChildrenTypes.All_Children);
                        List<int> test2 = new List<int>();
                        foreach (var abc in test)
                        {
                            test2.Add(abc.Index);
                        }
                        searchItems = searchItems.Except(test2).ToList();
                        cas = test2.Count;
                        errorcount += searchItems.Count();
                    }
                }
                else if (dir[blockname] == "Y+")
                {
                    path = -Ylen;
                    for (int i = 1; i < 5; i++)
                    {
                        List<float> box = new List<float>() { boxori[0], boxori[1] - Ylen * i / 4, boxori[2], boxori[3], boxori[4] - Ylen * i / 4, boxori[5] };
                        List<int> searchItems = Connector.GetObjectsInArea(box.ToArray(), new int[] { }, CrossBoundBox.IncludingPart);//IncludingPart
                        List<NodeVO> test = Connector.GetChildObjects(a.Index, ChildrenTypes.All_Children);
                        List<int> test2 = new List<int>();
                        foreach (var abc in test)
                        {
                            test2.Add(abc.Index);
                        }
                        searchItems = searchItems.Except(test2).ToList();
                        cas = test2.Count;
                        errorcount += searchItems.Count();
                    }
                }
                else
                {
                    path = Zlen;
                    for (int i = 1; i < 5; i++)
                    {
                        List<float> box = new List<float>() { boxori[0], boxori[1], boxori[2] + Zlen * i / 4, boxori[3], boxori[4], boxori[5] + Zlen * i / 4 };
                        List<int> searchItems = Connector.GetObjectsInArea(box.ToArray(), new int[] { }, CrossBoundBox.Fullycontained);//IncludingPart
                        List<NodeVO> test = Connector.GetChildObjects(a.Index, ChildrenTypes.All_Children);
                        List<int> test2 = new List<int>();
                        foreach(var abc in test)
                        {
                            test2.Add(abc.Index);
                        }
                        searchItems = searchItems.Except(test2).ToList();
                        cas = test2.Count;
                        errorcount += searchItems.Count();
                    }
                }
                File.AppendAllText(savepath,sw.Elapsed.ToString() + Environment.NewLine);
                File.AppendAllText(savepath, string.Format("{0}_{1}_{2}_{3}_오류정도 : {4}", blockname,errorcount,dir[blockname],cas,errorcount/cas ) + Environment.NewLine);
                if(cas==0)
                {
                    continue;
                }
                if (errorcount/cas>=10)
                {
                    Connector.ShowWaitDialog(false);
                    if(blockname=="2NDPE")
                    {
                        continue;
                    }
                    if(cas<10000)
                    {
                        continue;
                    }
                    MessageBox.Show(string.Format("\"{0}\"의\"{1}\"탑재방향이 잘못되었습니다.", blockname, dir[blockname]), "탑재방향 오류");
                    return false;
                }
            }
            sw.Stop();

            Connector.ShowWaitDialog(false);
            return true;
        }
        private void label18_Click(object sender, EventArgs e)
        {

        }
        List<string> Blocklist1 = new List<string>();
        List<string> Blocklist2 = new List<string>();
        List<string> Blocklistunit = new List<string>();
        private void button3_Click(object sender, EventArgs e)
        {
            allblock = true;
            Regex engRegex = new Regex(@"[a-zA-Z]");
            Regex numRegex = new Regex(@"[0-9]");

            Blocklist1.Clear(); Blocklist2.Clear(); Blocklistunit.Clear();

            Blocklist1 = (from blockdata in tree
                          where (blockdata.Split('/').Last().First() == '1' ||
                            blockdata.Split('/').Last().First() == '5' || blockdata.Split('/').Last().First() == '9') &&
                            blockdata.Split('/').Last().Length == 4
                          orderby blockdata.Split('/').Last()
                          select blockdata.Split('/').Last()).ToList().Distinct().ToList();
            Blocklist2 = (from blockdata in tree
                          where (blockdata.Split('/').Last().First() == '2' ||
                             blockdata.Split('/').Last().First() == '6') &&
                            blockdata.Split('/').Last().Length == 4
                          orderby blockdata.Split('/').Last()
                          select blockdata.Split('/').Last()).ToList().Distinct().ToList();
            Blocklistunit = (from blockdata in tree
                             where engRegex.IsMatch(blockdata.Split('/').Last().First().ToString()) && blockdata.Split('/').Last().Length == 4
                             && numRegex.IsMatch(blockdata.Split('/').Last()[1].ToString())
                             && numRegex.IsMatch(blockdata.Split('/').Last()[2].ToString()) && engRegex.IsMatch(blockdata.Split('/').Last().Last().ToString())
                             orderby blockdata.Split('/').Last()
                             select blockdata.Split('/').Last()).ToList().Distinct().ToList();



            Console.WriteLine("\n\n\n\n\nStart, cbBox1_SelectedIndexChanged");
            //

            tabControl1.SelectedTab = tabPage1;

            watListView1.Items.Clear();
            listView2.Items.Clear();
            if (cbBox1.SelectedItem.ToString() == "") { return; }

            //Console.WriteLine("_originPath:{0}", _originPath);
            _selectedSHIPPath = Path.Combine(_originPath, cbBox1.SelectedItem.ToString());
            _erectPath = Path.Combine(_originPath, "ERECT", cbBox1.SelectedItem.ToString());
            DirectoryInfo CC = new DirectoryInfo(_erectPath);
            if (!CC.Exists)
            {
                CC.Create();
            }

            //
            _shipNo = _selectedSHIPPath.Substring(_selectedSHIPPath.LastIndexOf(@"\") + 1);
            Console.WriteLine("_shipNo:{0}", _shipNo);
            proj = "AM121KBS";
            if (companyName == "HMD")
            {
                string projectDir = String.Format(@"Q:\project\{0}", _shipNo);
                Console.WriteLine("projectDir:{0}", projectDir);

                if (Directory.Exists(projectDir) == true)
                {
                    Console.WriteLine("Exist, projectDir:{0}", projectDir);

                    DirectoryInfo dir = new DirectoryInfo(projectDir);
                    DirectoryInfo[] subFiles = dir.GetDirectories();
                    foreach (DirectoryInfo ff in subFiles)
                    {
                        string checkString = ff.Name.ToString().ToUpper().Trim();
                        Console.WriteLine("checkString:{0}", checkString);

                        if (checkString.EndsWith("000") == true)
                        {
                            proj = checkString.Substring(0, checkString.LastIndexOf("000"));
                            break;
                        }
                    }
                }
            }
            Console.WriteLine("proj:{0}", proj);

            //
            _tempDir2 = Path.Combine(_tempDir, String.Format("{0}_{1}_{2}", _shipNo, myProc.Id, am12USERNAME));
            if (Directory.Exists(_tempDir2) == false)
            {
                try { Directory.CreateDirectory(_tempDir2); }
                catch { Console.WriteLine("Failed, Create _tempDir2:{0}", _tempDir2); }
            }
            //

            Console.WriteLine("_selectedSHIPPath:{0}, _shipNo:{1}", _selectedSHIPPath, _shipNo);

            if (String.IsNullOrEmpty(_selectedSHIPPath) == true) { return; }

            dic_treenode2Path.Clear();
            _treeAllFile = Path.Combine(_selectedSHIPPath, @"Z99_TREE_ALL.TXT");
            if (File.Exists(_treeAllFile) == true)
            {
                string[] readLines = File.ReadAllLines(_treeAllFile, Encoding.Default);
                tree = readLines.ToList();
                foreach (string readLine in readLines)
                {

                    if (readLine == "") { continue; }

                    string readLineReplace = Path.Combine(_originPath, String.Format("{0}", readLine.Replace(@"/", @"\").Substring(1)));
                    //Console.WriteLine("readLineReplace:{0}", readLineReplace);07

                    string nodeName = "";
                    try { nodeName = readLine.Substring(readLine.LastIndexOf(@"/") + 1).Trim(); }
                    catch { }

                    if (dic_treenode2Path.ContainsKey(nodeName) == true) { }
                    else { dic_treenode2Path[nodeName] = readLineReplace; }
                }
            }
            Console.WriteLine("dic_treenode2Path.Count:{0}", dic_treenode2Path.Count);

            //
            dic_Block2Date.Clear();
            List<string> erectlist = null;
            _erectDateFile = Path.Combine(_selectedSHIPPath, @"Z99_ERECT_ALL_NULL.TXT");
            if (File.Exists(_erectDateFile) == true)
            {
                erectlist = new List<string>();
                erectlist = ReadTxt(_erectDateFile);
            }
            if (erectlist != null)
            {
                foreach (string erectl in erectlist)
                {
                    string erectB = erectl.Split(',')[0].Trim();
                    string erectD = erectl.Split(',')[1].Trim();
                    if (!dic_Block2Date.ContainsKey(erectB))
                    {
                        dic_Block2Date[erectB] = erectD;
                    }
                }
            }
            Console.WriteLine("dic_Block2Date.Count:{0}", dic_Block2Date.Count);
            //

            //
            dic_Block2No.Clear(); dic_BlockDir.Clear();
            List<string> erectSavedlist = null;
            _erectSavedFile = Path.Combine(_erectPath, @"Z98_ERECT_ALLBLOCK_SAVED.TXT");
            if (File.Exists(_erectSavedFile) == true)
            {
                erectSavedlist = new List<string>();
                erectSavedlist = ReadTxt(_erectSavedFile);
            }
            if (erectSavedlist != null)
            {
                foreach (string erectSaved in erectSavedlist)
                {
                    if (erectSaved.StartsWith("#") == true)
                    {
                        try { _savedUser = erectSaved.Split(',')[1].Trim(); }
                        catch { }

                        continue;
                    }

                    string erectB = erectSaved.Split(',')[0].Trim();
                    string erectD = erectSaved.Split(',')[1].Trim();
                    if (erectSaved.Split(',').ToList().Count >= 3)
                    {
                        string erectDir = erectSaved.Split(',')[2].Trim();
                        if (!dic_BlockDir.ContainsKey(erectB))
                        {
                            dic_BlockDir.Add(erectB, erectDir);
                        }
                    }
                    if (!dic_Block2No.ContainsKey(erectB))
                    {
                        dic_Block2No.Add(erectB, erectD);
                    }
                }
            }
            Console.WriteLine("dic_Block2No.Count:{0}", dic_Block2No.Count);
            //

            label11.Text = String.Format("탑재일자 수정 ({0})", _savedUser);
            List<string> blda = new List<string>();
            blda.AddRange(Blocklist2); blda.AddRange(Blocklist1); blda.AddRange(Blocklistunit);

            ListViewUpdates(blda);

            //
            Console.WriteLine("End, cbBox1_SelectedIndexChanged");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if(cbBox1.Text=="")
            {
                MessageBox.Show("호선을 선택하세요");
                return;
            }
            ProgBar pb = new ProgBar(3, "탑재일자를 가져오고 있습니다....");
            pb.Show();
            RUNAP.RUN(cbBox1.Text, companyName, "REV");
            pb.Close();
            cbBox1_SelectedIndexChanged(null, null);
        }
        ListViewItem curItem;
        bool cancelEdit;
        ListViewItem.ListViewSubItem curSB;
        private void watListView1_MouseClick(object sender, MouseEventArgs e)
        {
            curItem = watListView1.GetItemAt(e.X, e.Y);
            if (curItem == null)
                return;

            curSB = curItem.GetSubItemAt(e.X, e.Y);
            int idxSub = curItem.SubItems.IndexOf(curSB);

            switch (idxSub)
            {
                case 8://5번째 Subitem만 수정가능하게
                    break;
                default:
                    return;

            }

            int ILeft = curSB.Bounds.Left + 2;
            int IWidrh = curSB.Bounds.Width;
            ChangeBOX.SetBounds(ILeft + watListView1.Left, curSB.Bounds.Top + watListView1.Top, IWidrh, curSB.Bounds.Height);

            ChangeBOX.Text = curSB.Text.Replace("mm","");
            ChangeBOX.Show();
            ChangeBOX.Focus();
        }

        private void ChangeBOX_KeyDown(object sender, KeyEventArgs e)
        {
            // 엔터키 수정 ESC키 취소
            switch (e.KeyCode)
            {
                case System.Windows.Forms.Keys.Enter:
                    cancelEdit = false;
                    e.Handled = true;
                    ChangeBOX.Hide();
                    break;
                case System.Windows.Forms.Keys.Escape:
                    cancelEdit = true;
                    e.Handled = true;
                    ChangeBOX.Hide();
                    break;
            }
        }

        private void ChangeBOX_Leave(object sender, EventArgs e)
        {
            ChangeBOX.Hide();
            if (cancelEdit == false)
            {
                //if (textBox1.Text.Trim() != "")
                //{
                curSB.Text = ChangeBOX.Text+"mm";

                int idxSub = curItem.SubItems.IndexOf(curSB);
                int idx = curItem.Index;

                Console.Write(curSB.Text);//
                //}
            }
            else
            {
                cancelEdit = false;
            }
            watListView1.Focus();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        bool keyin = false;
        private void cbBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                keyin = true;
                cbBox1_SelectedIndexChanged(null, null);
                keyin = false;
            }
        }    
    }
    public enum ClashStatus
    {
        NONE,
        TESTING,
        TESTED,
        SKIP
    }
    class TestDeepCopy
    {
        public int field1;
        public int field2;

        public TestDeepCopy DeepCopy()
        {
            TestDeepCopy copy = new TestDeepCopy();
            copy.field1 = this.field1;
            copy.field2 = this.field2;

            return copy;
        }

        public void Print()
        {
            Console.WriteLine("{0} {1}", field1, field2);
        }
    }

    public class ClashTaskMultiVO
    {
        public IVIZZARDService Connector { get; set; }
        public ClashTaskVO TaskVO { get; set; }

        public bool IsCompleted { get; set; }
        public bool IsTesting { get; set; }

        public ListViewItem ListItem { get; set; }

        public List<ClashResultVO> ResultItem { get; set; }

        public DateTime ClashTestStartDate { get; set; }
        public DateTime ClashTestFinishDate { get; set; }

        public string MOVING_BLOCK { get; set; }
        public string LUGWIRE { get; set; }
        public List<string> FIXED_BLOCK { get; set; }

        public ClashTaskMultiVO()
        {
            IsTesting = false;
            IsCompleted = false;


            ResultItem = new List<ClashResultVO>();
            FIXED_BLOCK = new List<string>();
        }

        public void StartClashTestDate()
        {
            ClashTestStartDate = DateTime.Now;

            if (ListItem != null)
                ListItem.BackColor = Color.Yellow;
        }

        public void FinishClashTestDate()
        {
            ClashTestFinishDate = DateTime.Now;

            if (ListItem != null)
                ListItem.BackColor = Color.FromKnownColor(KnownColor.Window);
        }

        public int GetElapsedSec()
        {
            TimeSpan ts = ClashTestFinishDate - ClashTestStartDate;

            int retInt = 0;
            try { retInt = Convert.ToInt32(ts.TotalSeconds); }
            catch { }

            return retInt;
        }
        public List<string> ExportResult99(string imagepath = "",CheckState imagecheck=CheckState.Unchecked) //CSV
        {      
            Console.WriteLine("1");
            //Console.WriteLine("Start, ExportResult99");
            List<string> items = new List<string>();

            string FIXED_BLOCK_STR = "";
            foreach (string FIXED_B in FIXED_BLOCK)
            {
                //Console.WriteLine("FIXED_B:{0}", FIXED_B);
                if (String.IsNullOrEmpty(FIXED_BLOCK_STR) == false) { FIXED_BLOCK_STR += "@"; }
                FIXED_BLOCK_STR += FIXED_B;
            }
            Console.WriteLine("2");
            //////Console.WriteLine("FIXED_BLOCK_STR:{0}", FIXED_BLOCK_STR);

            //PIPE-STRU 재 계산 Logic 적용
            bool check = false; string PIPE = "A";
            ObjectPropertyVO prop = new ObjectPropertyVO();
            Stopwatch sw = new Stopwatch();
            sw.Start();
            Console.WriteLine("3");
            Dictionary<int, double> recal = new Dictionary<int, double>();
            

            // 스냅샷 메모리 버퍼 설정
            Connector.SetMemBufferMode(300, 300);
            for (int i = 0; i < ResultItem.Count; i++)
            {
                ClashResultVO result = ResultItem[i];
                if (result.ResultType == ClashResultType.Unknown)
                    continue;

                double dis = result.Distance;

                items.Add(String.Format("{0}∫{1}∫{2}∫{3}∫{4}∫{5}∫{6}∫{7}∫{8}∫{9}∫{10}∫{11}∫{12}∫{13}∫{14}∫{15}∫{16}∫{17}∫{18}∫{19}∫{20}∫{21}∫{22}∫{23}∫{24}",
                    i + 1, //0
                    MOVING_BLOCK, //1
                    FIXED_BLOCK_STR, //2
                    result.PartNodeA, //3 (Model1)
                    result.PartNodeB, //4 (Modle2)
                    result.ResultType, //5
                    dis, //6F
                    result.PointX, //7
                    result.PointY, //8
                    result.PointZ, //9
                    "", //panelName1, //10
                    "", //panelName2, //11
                    TaskVO.CtType, //12
                    result.ClashResultString, //13
                    "", //modelDept1, //14
                    "", //modelDept2, //15
                    "", //modelType1, //16
                    "", //modelType2, //17
                    "", //18 O X
                    "", //sosokBLK, //19
                    result.NodePathA, //20
                    result.NodePathB, //21
                    "", //moduleORassy1, //22
                    "", //moduleORassy2, //23
                    "" //24
                    ));
                if (imagecheck == CheckState.Checked)
                {
                    ClashResultVO clashResultVO = ResultItem[i];

                    if (clashResultVO == null)
                        continue;

                    // 간섭 검사 결과 화면 설정
                    Connector.SetClashTestView(clashResultVO);

                    // 카메라 전환
                    Connector.SetCameraMode(CameraModes.ISOPLUS);

                    // 이미지 내보내기
                    string strPath = Connector.GetCurrentScreenToFileWithoutBuffer(300, 300, 100, false, false, 0);
                    DirectoryInfo newfile = new DirectoryInfo(imagepath);
                    if (!newfile.Exists)
                    {
                        newfile.Create();
                    }
                    string movePath = Path.Combine(imagepath, (i + 1).ToString() + ".jpg");
                    FileInfo existCheck = new FileInfo(movePath);
                    if (existCheck.Exists)
                    {
                        existCheck.Delete();
                    }
                    File.Move(strPath, movePath);
                }
            }
            // 스냅샷 메모리 버퍼 해제
            Connector.ReleaseMemBufferMode();
            return items;
        }

        /*bool IsEnglish(char ch)
        {

            if ((0x61 <= ch && ch <= 0x7A) || (0x41 <= ch && ch <= 0x5A))

                return true;

            else

                return false;

        }*/

        /*bool IsNumeric(char ch)
        {

            if ((0x61 <= ch && ch <= 0x7A) || (0x41 <= ch && ch <= 0x5A))

                return true;

            else

                return false;
        }*/

        bool IsNumericStr(string str)
        {
            bool isNumericStr = true;
            double dou = 0.0;
            try
            {
                dou = double.Parse(str);
                isNumericStr = true;
            }
            catch
            {
                isNumericStr = false;
            }
            return isNumericStr;
        }
    }
    class ListViewItemComparer : IComparer
    {
        bool first = true;
        private int col;
        public string sort = "asc";
        public ListViewItemComparer()
        {
            col = 0;
        }

        public ListViewItemComparer(int column, string sort)
        {
            col = column;
            this.sort = sort;
        }

        public int Compare(object x, object y)
        {
            if (sort == "asc")
                return String.Compare(((ListViewItem)x).SubItems[col].Text, ((ListViewItem)y).SubItems[col].Text);
            else
                return String.Compare(((ListViewItem)y).SubItems[col].Text, ((ListViewItem)x).SubItems[col].Text);
        }
    }

    class MyClass : ICloneable
    {

        public int MyField1;
        public int MyField2;

        public Object Clone()
        {
            MyClass newCopy = new MyClass();
            newCopy.MyField1 = this.MyField1;
            newCopy.MyField2 = this.MyField2;

            return newCopy;
        }
    }

    public class ClashData
    {
        public string BlockName { get; set; }
        public ClashStatus Status { get; set; }
        public ClashVO clash { get; set; }
        public string BlockPath { get; set; }
        public ClashData()
        {
            BlockName = string.Empty;
            BlockPath = string.Empty;
            Status = ClashStatus.NONE;
            clash = null;
        }
    }
}
