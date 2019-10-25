using System;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using MultiTaskServerInspection.Common;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Collections.ObjectModel;

namespace MultiTaskServerInspection
{
    public partial class InspectionForm : Form
    {
        //カテゴリ名
        string category_cpu = "Processor";
        string category_mem = "Memory";
        string category_network = "Network Interface";

        //カウンタ名
        string counter_cpu = "% Processor Time";
        string counter_mem = "Available MBytes";
        string counter_network = "Bytes Total/sec";

        //インスタンス名
        string instance_cpu = "_Total";
        string instanceMainDB_net = "ネットワーク機器名1";
        string instanceDC_net = "ネットワーク機器名2";
        string instanceSubDB_net = "ネットワーク機器名3";
        string instanceDevDB_net = "ネットワーク機器名4";
        string instanceFileServer1_net = "ネットワーク機器名5";
        string instanceFileServer2_net = "ネットワーク機器名6";
        string instanceFileServer3_net = "ネットワーク機器名7";

        //リモートコンピュータ名

        string DC = "DC";
        string MainDB = "MainDB";
        string SubDB = "SubDB";
        string DevDB = "DevDB";
        string FileServer1 = "FileServer1";
        string FileServer2 = "FileServer2";
        string FileServer3 = "FileServer3";

        PerformanceCounter UseRate;

        public InspectionForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                AsyncGetUserRate();
                FetchTabaleSpace();
                BatchFileInspec();
                //FetchEventlog();
                //EventExist();
            }
            catch(Exception ex)
            {
                String errorMesssage;
                errorMesssage = ex.Message + "\r\n\r\n" + ex.StackTrace;
                MessageBox.Show(errorMesssage, "例外エラー");
            }
        }
        /// <summary>
        /// リソース取得呼び出し元
        /// </summary>
        private async void AsyncGetUserRate()
        {
            try
            {
                //昨日のファイルの存在を感知した場合削除
                if (File.Exists(Application.StartupPath + @"\LOG\ResouceInfo.csv"))
                {
                    File.Delete(Application.StartupPath + @"\LOG\ResouceInfo.csv");
                }
                while (true)
                {
                    //cpu info
                    MainDBCPU.Text = await Task.Run(() => GetUseRate(category_cpu, counter_cpu, instance_cpu, MainDB));
                    SubDBCPU.Text = await Task.Run(() => GetUseRate(category_cpu, counter_cpu, instance_cpu, SubDB));
                    DevDBCPU.Text = await Task.Run(() => GetUseRate(category_cpu, counter_cpu, instance_cpu, DevDB));
                    DCCPU.Text = await Task.Run(() => GetUseRate(category_cpu, counter_cpu, instance_cpu, DC));
                    FileServer1CPU.Text = await Task.Run(() => GetUseRate(category_cpu, counter_cpu, instance_cpu, FileServer1));
                    FileServer2CPU.Text = await Task.Run(() => GetUseRate(category_cpu, counter_cpu, instance_cpu, FileServer2));
                    FileServer3CPU.Text = await Task.Run(() => GetUseRate(category_cpu, counter_cpu, instance_cpu, FileServer2));

                    //memory info
                    MainDBMEM.Text = await Task.Run(() => GetUseRate(category_mem, counter_mem, string.Empty, MainDB));
                    SubDBMEM.Text = await Task.Run(() => GetUseRate(category_mem, counter_mem, string.Empty, SubDB));
                    DevDBMEM.Text = await Task.Run(() => GetUseRate(category_mem, counter_mem, string.Empty, DevDB));
                    DCMEM.Text = await Task.Run(() => GetUseRate(category_mem, counter_mem, string.Empty, DC));
                    FileServer1MEM.Text = await Task.Run(() => GetUseRate(category_mem, counter_mem, string.Empty, FileServer1));
                    FileServer2MEM.Text = await Task.Run(() => GetUseRate(category_mem, counter_mem, string.Empty, FileServer2));
                    FileServer3MEM.Text = await Task.Run(() => GetUseRate(category_mem, counter_mem, string.Empty, FileServer3));
                    
                    ResouceInfoToCSV
                        (
                        MainDBCPU.Text,
                        SubDBCPU.Text,
                        DevDBCPU.Text,
                        DCCPU.Text,
                        FileServer1CPU.Text,
                        FileServer2CPU.Text,
                        FileServer3CPU.Text,
                        MainDBMEM.Text,
                        SubDBMEM.Text,
                        DecDBMEM.Text,
                        DCMEM.Text,
                        FileServer1MEM.Text,
                        FileServer2MEM.Text,
                        FileServer3MEM.Text
                        );
                    
                    //Network info
                    MainDBNET.Text = await Task.Run(() => GetUseRate(category_network, counter_network, instanceDC_net, MainDB));
                    SubDBNET.Text = await Task.Run(() => GetUseRate(category_network, counter_network, instanceSubDB_net, SubDB));
                    DevDBNET.Text = await Task.Run(() => GetUseRate(category_network, counter_network, instanceDevDB_net, DevDB));
                    DCNET.Text = await Task.Run(() => GetUseRate(category_network, counter_network, instanceDC_net, DC));
                    FileServer1NET.Text = await Task.Run(() => GetUseRate(category_network, counter_network, instanceDC_net, FileServer1));
                    FileServer2NET.Text = await Task.Run(() => GetUseRate(category_network, counter_network, instanceFileServer1_net, FileServer2));
                    FileServer3NET.Text = await Task.Run(() => GetUseRate(category_network, counter_network, instanceFileServer2_net, FileServer3));

                    await Task.Delay(1000);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// リソース取得メソッド
        /// </summary>
        /// <param name="categoryName">カテゴリ名</param>
        /// <param name="counterName">カウンタ名</param>
        /// <param name="instanceName">インスタンス名</param>
        /// <param name="machineName">リモートコンピュータ名</param>
        private string GetUseRate(string categoryName, string counterName, string instanceName, string machineName)
        {
            try
            {
                string UserRate = string.Empty;
                UseRate = new PerformanceCounter(categoryName, counterName, instanceName, machineName);
                if (categoryName == "Processor")
                {
                    UseRate.NextValue().ToString();
                    //UserRate = UseRate.NextValue().ToString("F2") + "  %";
                    UserRate = UseRate.NextValue().ToString("F2");
                }
                else if (categoryName == "Memory")
                {
                    UseRate.NextValue().ToString();
                    //UserRate = (UseRate.NextValue() / 1000).ToString("F2") + "  GB";
                    UserRate = (UseRate.NextValue() / 1000).ToString("F2");
                }
                else
                {
                    UseRate.NextValue().ToString();
                    //UserRate = (UseRate.NextValue() / Math.Pow(1000, 2)).ToString("F2") + "  MBytes /sec";
                    UserRate = (UseRate.NextValue() / Math.Pow(1000, 2)).ToString("F2");
                }
                return UserRate;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// 表領域情報取得メソッド
        /// </summary>
        private  void FetchTabaleSpace()
        {
            try
            {
                DataTable MainDB = new DataTable();
                DataTable SubDB = new DataTable();
                DataTable DevDB = new DataTable();
                DBRS dr = new DBRS();

                MainDB = dr.TablespaceStatement("Main", "XXX", "MainIP");
                // 表領域をそれぞれ取得
                var row1 = MainDB.AsEnumerable().Where(d => d.Field<string>("TABLESPACE_NAME") == ("Main")).ToArray();
                MainSpace.Text = row1[0]["free_GByte"].ToString() + (" / ").ToString() + row1[0]["Total_Gbyte"].ToString() + " GB".ToString();
                row1 = MainDB.AsEnumerable().Where(d => d.Field<string>("TABLESPACE_NAME") == ("SYSAUX")).ToArray();
                SubSpace.Text = row1[0]["free_GByte"].ToString() + (" / ").ToString() + row1[0]["Total_Gbyte"].ToString() + " GB".ToString();
                row1 = MainDB.AsEnumerable().Where(d => d.Field<string>("TABLESPACE_NAME") == ("SYSTEM")).ToArray();
                DevSpace.Text = row1[0]["free_GByte"].ToString() + (" / ").ToString() + row1[0]["Total_Gbyte"].ToString() + " GB".ToString();

                SubDB = dr.TablespaceStatement("Sub", "XXX", "SubIP");

                var row2 = SubDB.AsEnumerable().Where(d => d.Field<string>("TABLESPACE_NAME") == ("Sub")).ToArray();
                SubSpace.Text = row2[0]["free_GByte"].ToString() + (" / ").ToString() + row2[0]["Total_Gbyte"].ToString() + " GB".ToString();
                row2 = SubDB.AsEnumerable().Where(d => d.Field<string>("TABLESPACE_NAME") == ("SYSAUX")).ToArray();
                SYSAUXSPACE012.Text = row2[0]["free_GByte"].ToString() + (" / ").ToString() + row2[0]["Total_Gbyte"].ToString() + " GB".ToString();
                row2 = SubDB.AsEnumerable().Where(d => d.Field<string>("TABLESPACE_NAME") == ("SYSTEM")).ToArray();
                SYSTEMSPACE012.Text = row2[0]["free_GByte"].ToString() + (" / ").ToString() + row2[0]["Total_Gbyte"].ToString() + " GB".ToString();

                DevDB = dr.TablespaceStatement("Dev", "XXX", "DevIP");

                var row3 = DevDB.AsEnumerable().Where(d => d.Field<string>("TABLESPACE_NAME") == ("Dev")).ToArray();
                DevSpace.Text = row3[0]["free_GByte"].ToString() + (" / ").ToString() + row3[0]["Total_Gbyte"].ToString() + " GB".ToString();
                row3 = DevDB.AsEnumerable().Where(d => d.Field<string>("TABLESPACE_NAME") == ("SYSAUX")).ToArray();
                SYSAUXSPACE021.Text = row3[0]["free_GByte"].ToString() + (" / ").ToString() + row3[0]["Total_Gbyte"].ToString() + " GB".ToString();
                row3 = DevDB.AsEnumerable().Where(d => d.Field<string>("TABLESPACE_NAME") == ("SYSTEM")).ToArray();
                SYSTEMSPACE021.Text = row3[0]["free_GByte"].ToString() + (" / ").ToString() + row3[0]["Total_Gbyte"].ToString() + " GB".ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            } 
        }
        /// <summary>
        /// ダンプファイル点検メソッド
        /// </summary>
        private  void BatchFileInspec()
        {
            
            try
            {
                string[] Main = new string[8];
                Array.Copy(DmpFileInfo(@"\\instanceName1\g$\ORAEX"), Main, Main.Length);
                MainDBexpStartTime.Text = Main[0];
                MainDBexpError.Text = Main[2];
                MainDBexpEndTime.Text = Main[1];
                MainDBexpFileSize.Text = Main[3] != string.Empty ? String.Format("{0:#,0 }KB",Int64.Parse(Main[3]) / 1000) : string.Empty;
                MainDBimpStartTime.Text = Main[5];
                MainDBimpError.Text = Main[7];
                MainDBimpEndTime.Text = Main[6];
                MainDBimpFileSize.Text = Main[4] != string.Empty ? String.Format("{0:#,0 }KB", Int64.Parse(Main[4]) / 1000) : string.Empty;

                string[] Sub = new string[8];
                Array.Copy(DmpFileInfo(@"\\instanceName2\g$\ORAEX"), Sub, Sub.Length);
                SubDBexpStartTime.Text = Sub[0];
                SubDBexpError.Text = Sub[2];
                SubDBexpEndTime.Text = Sub[1];
                SubDBexpFileSize.Text = Sub[3] != string.Empty ? String.Format("{0:#,0 }KB", Int64.Parse(Sub[3]) / 1000) : string.Empty;
                SubDBimpStartTime.Text = Sub[5];
                SubDBimpError.Text = Sub[7];
                SubDBimpEndTime.Text = Sub[6];
                SubDBimpFileSize.Text = Sub[4] != string.Empty ? String.Format("{0:#,0 }KB", Int64.Parse(Sub[4]) / 1000) : string.Empty;

                string[] Dev = new string[8];
                Array.Copy(DmpFileInfo(@"\\instanceName3\g$\ORAEX"), Dev, Dev.Length);
                DevDBexpStartTime.Text = Dev[0];
                DevDBexpError.Text = Dev[2];
                DevDBexpEndTime.Text = Dev[1];
                DevDBexpFileSize.Text = Dev[3] != string.Empty ? String.Format("{0:#,0 }KB", Int64.Parse(Dev[3]) / 1000) : string.Empty;
                DevDBimpStartTime.Text = Dev[5];
                DevDBimpError.Text = Dev[7];
                DevDBimpEndTime.Text = Dev[6];
                DevDBimpFileSize.Text = Dev[4] != string.Empty ? String.Format("{0:#,0 }KB", Int64.Parse(Dev[4]) / 1000) : string.Empty;

            }
            catch (Exception ex)
            {
                throw ex;
            }
                
        }
        
        /// <summary>
        ///  副番サーバー専用 点検メソッド
        /// </summary>
        /// <param name="FilePath">対象ファイルの階層ディレクトリ</param>
        /// <returns></returns>
        private string[] DmpFileInfo(string FilePath)
        {
            string[] FileLogInfo = new string[8] 
            {
                /*
                 exp開始時間
                 exp終了時間
                 experr check
                 expdmp容量
                 imp開始時間
                 imp終了時間
                 imperr check
                 */
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty
            };
            FileStream fs = null;
            StreamReader sr = null;
            try
            {
                int i = 0;
                string impDmp = FilePath + @"\IMP\DB-BAK.DMP";
                string impLog = FilePath + @"\LOG\DB-BAK-imp.log";
                string expDmp = FilePath + @"\EXP\DB-BAK.DMP";
                string expLog = FilePath + @"\LOG\DB-BAK-exp.LOG";
                string err = string.Empty;
                string line = string.Empty;
                FileInfo impfi = new FileInfo(impDmp);
                FileInfo expfi = new FileInfo(expDmp);
                string DBNAME = expLog.Substring(2, 8).ToUpper(); //データベースの名前
                CommonMethod omMeth = new CommonMethod();

                if (omMeth.IsDmpCheck(DBNAME, 1) == true)
                {
                    if ((File.Exists(expLog)) && (File.GetLastWriteTime(expLog).Date == DateTime.Now.Date))
                    {
                        //FileStreamでのロックチェック
                        if (omMeth.IsFileLocked(expLog) == true)
                        {
                            err += DBNAME + Environment.NewLine + Path.GetFileName(expLog) + "がロックされています。\n";
                        }
                        else
                        {
                            fs = new FileStream(expLog, FileMode.Open, FileAccess.Read, FileShare.Read);
                            sr = new StreamReader(fs, Encoding.GetEncoding("shift_jis"));
                            while (sr.EndOfStream == false)
                            {
                                line = sr.ReadLine();
                                if (Regex.IsMatch(line, @"\d{1,2}月\s\d{1,2}\s\d\d:\d\d:\d\d\s\d{4,4}"))
                                {
                                    FileLogInfo[i] = (Regex.Match(
                                        Regex.Match(
                                            line, @"\d{1,2}月\s\d{1,2}\s\d\d:\d\d:\d\d\s\d{4,4}").ToString(),
                                            @"\d\d:\d\d:\d\d").ToString()
                                            );
                                    i++;
                                }

                                if (Regex.IsMatch(line, "で正常に完了しました"))
                                {
                                    FileLogInfo[2] = "エラー無";
                                }
                                else
                                {
                                    FileLogInfo[2] = "エラー有";
                                }
                            }
                        }
                    }
                    else
                    {
                        err += DBNAME + Environment.NewLine + Path.GetFileName(expLog) + "が存在しません。\n";
                    }
                }

                if (omMeth.IsDmpCheck(DBNAME, 2) == true)
                {
                     if ((File.Exists(expDmp)) && (File.GetLastWriteTime(expDmp).Date == DateTime.Now.Date))
                    {
                        if (omMeth.IsFileLocked(expDmp) == true)
                        {
                            err += DBNAME + Environment.NewLine + Path.GetFileName(expDmp) + "がロックされています。\n";
                        }
                        else
                        {
                            FileLogInfo[3] = expfi.Length.ToString();
                        }
                    }
                    else
                    {
                        err += DBNAME + Environment.NewLine + Path.GetFileName(expDmp) + "が存在しません。\n";
                    }
                }

                if (omMeth.IsDmpCheck(DBNAME, 3) == true)
                {
                    if ((File.Exists(impDmp)) && (File.GetLastWriteTime(impDmp).Date == DateTime.Now.Date))
                    {
                        if (omMeth.IsFileLocked(impDmp) == true)
                        {
                            err += Path.GetFileName(impDmp) + "がロックされています。\n";
                        }
                        else
                        {
                            FileLogInfo[4] = impfi.Length.ToString();
                        }
                    }
                    else
                    {
                        err += DBNAME + Environment.NewLine + Path.GetFileName(impDmp) + "が存在しません\n";
                    }
                }

                if (omMeth.IsDmpCheck(DBNAME, 4) == true)
                {
                    if ((File.Exists(impLog)) && (File.GetLastWriteTime(impLog).Date == DateTime.Now.Date))
                    {
                        //FileStreamでのロックチェック
                        if (omMeth.IsFileLocked(impLog) == true)
                        {
                            err += DBNAME + Environment.NewLine + Path.GetFileName(impLog) + "がロックされています。";
                        }
                        else
                        {
                            fs = new FileStream(impLog, FileMode.Open, FileAccess.Read, FileShare.Read);
                            sr = new StreamReader(fs, Encoding.GetEncoding("shift_jis"));
                            i = 5;// インポートの開始時間を格納する
                            while (sr.EndOfStream == false)
                            {
                                line = sr.ReadLine();
                                if (Regex.IsMatch(line, @"\d{1,2}月\s\d{1,2}\s\d\d:\d\d:\d\d\s\d{4,4}"))
                                {
                                    FileLogInfo[i] = (Regex.Match(
                                        Regex.Match(
                                            line, @"\d{1,2}月\s\d{1,2}\s\d\d:\d\d:\d\d\s\d{4,4}").ToString(),
                                            @"\d\d:\d\d:\d\d").ToString());
                                    i++;
                                }

                                if (Regex.IsMatch(line, "で正常に完了しました"))
                                {
                                    FileLogInfo[7] = "エラー無";
                                }
                                else
                                {
                                    FileLogInfo[7] = "エラー有";
                                }
                            }
                        }
                    }
                    else
                    {
                        err += DBNAME + Environment.NewLine + Path.GetFileName(impLog) + "が存在しません。\n";
                    }
                }
                if (err != string.Empty)
                {
                    if (err != string.Empty)
                    {
                        MessageBox.Show(err,
                            "エラー",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                            );
                    }
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }

                if (sr != null)
                {
                    sr.Close();
                }
            }
            return FileLogInfo;
        }

        private void EventExist()
        {

            string[] eLog = new string[7]
            {
                Application.StartupPath + @"\LOG\MainDB.csv",
                Application.StartupPath + @"\LOG\SubDB.csv",
                Application.StartupPath + @"\LOG\DevDB.csv",
                Application.StartupPath + @"\LOG\DC.csv",
                Application.StartupPath + @"\LOG\FileServer1.csv",
                Application.StartupPath + @"\LOG\FileServer2.csv",
                Application.StartupPath + @"\LOG\FileServer3.csv",
            };

            string[] eLogResult = new string[7]
            {
                "ExceptionErr",
                "ExceptionErr",
                "ExceptionErr",
                "ExceptionErr",
                "ExceptionErr",
                "ExceptionErr",
                "ExceptionErr"
            };

            CommonMethod omMeth = new CommonMethod();
            for (int e = 0;  e < eLog.Length; e++)
            {
                if (omMeth.IsExistEventLog(eLog[e]))
                {
                    eLogResult[e] = "エラー有";
                }
                else
                {
                    eLogResult[e] = "エラー無";
                }
            }
            MainEventViewer.Text = eLogResult[0];
            SubEventViewer.Text = eLogResult[1];
            DevEventViewer.Text = eLogResult[2];
            DCEventViewer.Text = eLogResult[3];
            FileServer1EventViewer.Text = eLogResult[4];
            FileServer2EventViewer.Text = eLogResult[5];
            FileServer3EventViewer.Text = eLogResult[6];
        }

        private void ResouceInfoToCSV
            (
            string MainCPU, string SubCPU, string DevCPU, string DCCPU, string FileServer1CPU, string FileServer2CPU, string FileServer3CPU,
            string MainMEM, string SubDBMEM, string DevMEM, string DCMEM, string FileServer1MEM, string FileServer2MEM, string FileServer3MEM
            )
        {
            List<string> ResouceInfo = new List<string>()
            {
                MainCPU, SubCPU, DevCPU, DCCPU, FileServer1CPU, FileServer2CPU, FileServer3CPU,
                MainMEM, SubDBMEM, DevMEM, DCMEM, FileServer1MEM, FileServer2MEM, FileServer3MEM
            };
            //Listの中身を全て数値に変換
            List<double> convertIntList = ResouceInfo.ConvertAll(double.Parse);

            StringBuilder sb = new StringBuilder();
            if (!File.Exists(Application.StartupPath + @"\LOG\ResouceInfo.csv"))
            {
                sb.AppendLine
                (
                "MainDBCPU, SubDBCPU, DevDBCPU, DCCPU, FileServer1CPU, FileServer2CPU, FileServer3CPU, " +
                "MainDBMEM, SubDBMEM, DecDBMEM, DCMEM, FileServer1MEM, FileServer2MEM, FileServer3MEM, SYSDATE"
                );
            }
            if (convertIntList.Any(x => x > 30.0))
            {
                sb.AppendLine
               (
                ResouceInfo[0] + " %" + "," + ResouceInfo[1] + " %" + "," + ResouceInfo[2] + " %" + "," + ResouceInfo[3] + " %" + "," + ResouceInfo[4] + " %" + "," + ResouceInfo[5] + " %" + "," + ResouceInfo[6] + " %" + "," +
                ResouceInfo[7] + " GB" + "," + ResouceInfo[8] + " GB" + "," + ResouceInfo[9] + " GB" + "," + ResouceInfo[10] + " GB" + "," + ResouceInfo[11] + " GB" + "," + ResouceInfo[12] + " GB" + "," + ResouceInfo[13] + " GB" + "," + DateTime.Now
               );
            }
            File.AppendAllText(Application.StartupPath + @"\LOG\ResouceInfo.csv", sb.ToString());
        }

        private void FetchEventlog()
        {
            try
            {
                string Path = Application.StartupPath + @"\LOG\FetchEventLog.ps1";

                RunspaceInvoke invoke = new RunspaceInvoke();
                Collection<PSObject> result = invoke.Invoke(Path);
                foreach (PSObject result_str in result)
                {
                    MessageBox.Show(result_str.ToString());
                    MessageBox.Show(Path);
                }
                invoke.Dispose();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }  

        private void EventViewer_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                string EventLogPath = Application.StartupPath + @"\LOG\MainDB.csv";
                CommonMethod omMeth = new CommonMethod();
                if (omMeth.IsExistEventLog(EventLogPath))
                {
                    MessageBox.Show("エラーです");
                }
            }
            catch (Exception ex)
            {
                String errorMesssage;
                errorMesssage = ex.Message + "\r\n\r\n" + ex.StackTrace;
                MessageBox.Show(errorMesssage, "例外エラー");
            }
        }
    }
}
