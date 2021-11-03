using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace autoCLASH.Lib
{
    class RUNAP
    {
        static string _tempDir2 = @"C:\Temp";
        static Process myProc = Process.GetCurrentProcess();
        static string am12EXEDir = @"C:\AVEVA\Marine\OH12.1.SP4";

        static string proj = "27256";
        static string am12USERNAME = "D337935";
        static string am12PASSWORD = "HMD337935";
        static string am12MDB = "/HULL";
        static string originPath = "";
        static string LogPath = "";

        static List<List<string>> _exceptBlock = new List<List<string>>();

        static int _WAITTIME = 10;

        static string companyName = "HMD";
        /// <summary>
        /// autoAP실행 구문
        /// </summary>
        /// <param name="shipno">호선명</param>
        /// <param name="company">회사명(HMD,HHISS)</param>
        /// <param name="sREVRVM">REVorRVM</param>
        public static void RUN(string shipno, string company, string sREVRVM)
        {
            proj = shipno;
            ArrayList writeLines = new ArrayList();

            if (company == "HMD")
            {
                originPath = @"\\210.118.131.6\simulation\__Simulation_Program_Server\AMPROJ_REV";
                LogPath = @"Q:\env\Vitesse\Infoget\AutoRunforClashCheck\RUN_LOG";
            }
            else if (company == "HHISS")
            {
                originPath = "??";
                LogPath = "??";
            }
            else
            {
                originPath = @"D:\autoCCS\Ship";
                LogPath = @"D:\autoCCS\Ship";
            }
            string projectDir = Path.Combine(originPath, proj);

            int iBlockSite = 20;
            int iZone = 200;
            int iRev = 10;

            int iDo = 1;

            string listsFile = Path.Combine(_tempDir2, String.Format(@"listsFile_{0}.txt", myProc.Id));
            FileClear(listsFile);
            writeLines.Clear();

            ////////////////////////////////////////////////////////////////////////////////////////// 1

            am12RunBAT("STEP1", listsFile); //Only Run

            ////////////////////////////////////////////////////////////////////////////////////////// 2

        }
        public static void FileClear(string listsFile)
        {
            FileInfo sw99 = new FileInfo(listsFile);
            StreamWriter writer99 = sw99.CreateText();
            writer99.Close();
        }

        public static List<string> GetblockList(List<string> z99file)
        {
            List<string> blocks = new List<string>();
            foreach (var a in z99file)
            {
                MatchCollection matches = Regex.Matches(a, "/");
                int cnt = matches.Count;
                if (cnt == 2)
                {
                    List<string> split = a.Split('/').ToList();
                    try
                    {
                        blocks.Add(split[2]);
                    }
                    catch { }
                }
            }
            return blocks;
        }
        public static void am12RunBAT(string sKind, string listsFile)
        {
            string pmlFile = Path.Combine(_tempDir2, String.Format(@"macrorun_{0}.pml", myProc.Id));
            Console.WriteLine("pmlFile:{0}", pmlFile);
            FileInfo sw2 = new FileInfo(pmlFile);
            StreamWriter writer2 = sw2.CreateText();

            writer2.WriteLine(String.Format(@"import 'MarAPI'"));
            writer2.WriteLine(String.Format(@"Handle any"));
            writer2.WriteLine(String.Format(@"endhandle"));
            writer2.WriteLine(String.Format(@"using namespace 'Aveva.Marine.Utility'"));
            writer2.WriteLine(String.Format(@"!!MUtil = object MarUtil()"));

            writer2.WriteLine(String.Format(@""));
            writer2.WriteLine(String.Format(@"$P ################# 0"));
            writer2.WriteLine(String.Format(@""));

            if (sKind == "STEP1")
            {
                writer2.WriteLine(String.Format(@"!!Mutil.tbenvironmentset('AUTOAP_EMPTY','{0}')", listsFile));
                writer2.WriteLine(String.Format(@"!!Mutil.tbenvironmentset('AUTOAP_BLOCKSITE','{0}')", "_"));
                writer2.WriteLine(String.Format(@"!!Mutil.tbenvironmentset('AUTOAP_ZONE','{0}')", "_"));
                writer2.WriteLine(String.Format(@"!!Mutil.tbenvironmentset('AUTOAP_REV','{0}')", "_"));
            }
            else if (sKind == "STEP2")
            {
                writer2.WriteLine(String.Format(@"!!Mutil.tbenvironmentset('AUTOAP_EMPTY','{0}')", "_"));
                writer2.WriteLine(String.Format(@"!!Mutil.tbenvironmentset('AUTOAP_BLOCKSITE','{0}')", listsFile));
                writer2.WriteLine(String.Format(@"!!Mutil.tbenvironmentset('AUTOAP_ZONE','{0}')", "_"));
                writer2.WriteLine(String.Format(@"!!Mutil.tbenvironmentset('AUTOAP_REV','{0}')", "_"));
            }
            else if (sKind == "STEP3")
            {
                writer2.WriteLine(String.Format(@"!!Mutil.tbenvironmentset('AUTOAP_EMPTY','{0}')", "_"));
                writer2.WriteLine(String.Format(@"!!Mutil.tbenvironmentset('AUTOAP_BLOCKSITE','{0}')", "_"));
                writer2.WriteLine(String.Format(@"!!Mutil.tbenvironmentset('AUTOAP_ZONE','{0}')", listsFile));
                writer2.WriteLine(String.Format(@"!!Mutil.tbenvironmentset('AUTOAP_REV','{0}')", "_"));
            }
            else if (sKind == "STEP4")
            {
                writer2.WriteLine(String.Format(@"!!Mutil.tbenvironmentset('AUTOAP_EMPTY','{0}')", "_"));
                writer2.WriteLine(String.Format(@"!!Mutil.tbenvironmentset('AUTOAP_BLOCKSITE','{0}')", "_"));
                writer2.WriteLine(String.Format(@"!!Mutil.tbenvironmentset('AUTOAP_ZONE','{0}')", "_"));
                writer2.WriteLine(String.Format(@"!!Mutil.tbenvironmentset('AUTOAP_REV','{0}')", listsFile));
            }

            writer2.WriteLine(String.Format(@""));
            writer2.WriteLine(String.Format(@"$P ################# 1"));
            writer2.WriteLine(String.Format(@""));

            writer2.WriteLine(String.Format(@"import 'autoAP'"));
            writer2.WriteLine(String.Format(@"Handle any"));
            writer2.WriteLine(String.Format(@"endhandle"));
            writer2.WriteLine(String.Format(@"using namespace 'autoAP'"));
            writer2.WriteLine(String.Format(@"!kkk2 = Object MainForm()"));
            writer2.WriteLine(String.Format(@"!kkk2.Run()")); //Run or Run2

            writer2.WriteLine(String.Format(@""));
            writer2.WriteLine(String.Format(@"$P ################# 2"));
            writer2.WriteLine(String.Format(@""));

            writer2.WriteLine(String.Format(@"FINISH"));

            writer2.Close();
            //

            string anyStr2 = Path.Combine(_tempDir2, String.Format(@"autoAP_{0}.txt", myProc.Id));
            string batchName = Path.Combine(_tempDir2, string.Format("autoAP_{0}.bat", myProc.Id)); //배치 런.
            if (companyName == "HMD")
            {
                AMPreRun(batchName, "", "RUN_MAC", "marhdes.exe", 0 + 1, pmlFile, anyStr2);
            }
            if (companyName == "HHISS")
            {
                //
            }
            string vbs = CreateVBSFile(batchName);

            string TempPath = @"C:\Temp";
            System.Diagnostics.Process ps = new System.Diagnostics.Process();
            ps.StartInfo.FileName = Path.Combine(_tempDir2, vbs);
            ps.StartInfo.WorkingDirectory = TempPath;
            ps.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;

            //위와 같이 StartInfo에 실행할 프로그램의 정보를 설정한 후 Start()를 실행하면 된다.
            ps.Start();
            int waittime = _WAITTIME * 60 * 1000;//분단위
            bool iswait = ps.WaitForExit(waittime);//프로세서가 끝 나기를 기다림 1초(1000밀리) 1시간

            //프로세스 강제로 죽이기
            try
            {
                Process[] pss = Process.GetProcessesByName("marhdes");
                if (pss.Length > 0)
                {
                    foreach (var a in pss)
                    {
                        a.Kill();
                    }
                }
                pss = Process.GetProcesses();
                if (pss.Length > 0)
                {
                    foreach (var a in pss)
                    {
                        if (a.ProcessName.Contains("PDMSConsole"))
                        {
                            a.Kill();
                        }
                    }
                }
            }
            catch { }

            Console.WriteLine("End, button1_Click");
        }
        public static string CreateVBSFile(string batchFilePath)
        {
            Console.WriteLine("Start, CreateVBSFile");

            string fileName = string.Format("{0}.vbs", Path.GetFileNameWithoutExtension(batchFilePath));
            string path = Path.Combine(Path.GetDirectoryName(batchFilePath), fileName); // vbs파일의 경로

            FileInfo sw = new FileInfo(path);
            StreamWriter writer = sw.CreateText();

            writer.WriteLine(String.Format("Set WshShell = WScript.CreateObject(\"WScript.Shell\")"));
            writer.WriteLine(String.Format("wshShell.run \"{0}\",0,True", batchFilePath));
            writer.WriteLine(String.Format("Set WshShell = nothing"));
            writer.Close();

            Console.WriteLine("End, CreateVBSFile");

            return path;
        }
        public static bool AMPreRun(string batchFile, string outputFile, string whatToDo, string module, int runTime, string anyStr, string anyStr2) //runTime -> File Split(Batch Run 분할) 하는 경우 대비.
        {
            Console.WriteLine("Start, AMPreRun");

            bool bAMPreRun = false;

            if (whatToDo == "RUN_MAC") //pml Run
            {
                FileInfo sw = new FileInfo(batchFile);
                StreamWriter writer = sw.CreateText();

                writer.WriteLine(string.Format("set PDMS_INSTALLED_DIR={0}", am12EXEDir));
                writer.WriteLine(string.Format("set PDMSEXE=%PDMS_INSTALLED_DIR%\\"));
                writer.WriteLine(string.Format("call %PDMS_INSTALLED_DIR%\\marstart.bat %PDMS_INSTALLED_DIR%"));
                writer.WriteLine(string.Format("call %PDMS_INSTALLED_DIR%\\evars.bat  %PDMS_INSTALLED_DIR%"));
                writer.WriteLine(string.Format("set UCase=ABCDEFGHIJKLMNOPQRSTUVWXYZ@"));
                writer.WriteLine(string.Format("set LCase=abcdefghijklmnopqrstuvwxyz@"));
                writer.WriteLine(string.Format("for /L %%A in (0,1,25) do Call :ToUpper !LCase:~%%A,1! !UCase:~%%A,1!"));
                writer.WriteLine(string.Format("set MDB=%MDBstring%"));
                writer.WriteLine(string.Format("set PROJ=%PROJstring%"));
                writer.WriteLine(string.Format("set CADC_LANG=KOREAN"));
                writer.WriteLine(string.Format("echo/"));

                writer.WriteLine(string.Format(@"{0} -TTY -proj={1} -user={2} -pass={3} -mdb={4} -appl=structural -macro=$M/{5} -console",
                 Path.Combine(am12EXEDir, "marhdes.exe"), proj, am12USERNAME, am12PASSWORD, am12MDB, anyStr)); //

                writer.Close();
            }

            Console.WriteLine("End, AMPreRun");

            return bAMPreRun;
        }
        private static void writeFile999(ArrayList lineArrayLists, string fileName, bool bANSI)
        {
            //File Write (ANSI or UTF-8 가능)
            FileStream fileStreamOutputXXX = new FileStream(fileName, FileMode.Create);
            fileStreamOutputXXX.Seek(0, SeekOrigin.Begin);

            foreach (string wrieLine in lineArrayLists)
            {
                byte[] info = System.Text.Encoding.Default.GetBytes(wrieLine + "\r\n"); //ANSI
                if (bANSI == false)
                {
                    info = new UTF8Encoding(true).GetBytes(wrieLine + "\r\n"); //UTF-8
                }
                fileStreamOutputXXX.Write(info, 0, info.Length);
            }

            fileStreamOutputXXX.Flush();
            fileStreamOutputXXX.Close();
            //
        }
    }
}
