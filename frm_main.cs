using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MAUTO
{
    public partial class frm_main : Form
    {
        public const string STR_SYS_PRE = "(Preparation)";
        public const string STR_STS_NOT_START = "";
        public const string STR_STS_EXECUTING = "EXECUTING";
        public const string STR_STS_DONE = "DONE";

        public string pStrConnecting;
        public string pStrMdbPath;

        public struct typ_task_status
        {
            public string seq_no;           // シーケンスID
            public string task_day;         // タスク実行日
            public string task_id;          // タスク実行ID
            public string task_status;      // タスク実行ステータス
            public string task_start;       // タスク実行開始時刻
            public string task_end;         // タスク実行終了時刻
            public string task_exe;         // タスク実行ファイル
            public string task_comment;     // タスク実行ファイル
        }
        public Dictionary<string, typ_task_status> pTaskStauts;

        public string strExecTaskId;        // 実行中タスク実行ID

        public frm_main()
        {
            InitializeComponent();
            // 初期化処理
            CommonLogger.pComLogFlag = true;

            this.pStrMdbPath = CommonApp.GetIniValue("MAUTO", "MDB_FILE", CommonApp.CON_INI_FILE);

            this.pStrConnecting = "";
            this.pStrConnecting = this.pStrConnecting + "Provider=" + CommonApp.GetIniValue("MAUTO", "MDB_PRV", CommonApp.CON_INI_FILE) + "; ";
            this.pStrConnecting = this.pStrConnecting + "Data Source =" + this.pStrMdbPath + "; ";
        }

        private void frm_main_Load(object sender, EventArgs e)
        {           
            string strExecTaskSeqNo;
            string strTaskExecTimeStamp;
            string strExecFile;
            string strMsg;

            try
            {
                // ***** タスクステータス取得 *****
                fnc_GetTaskList();

                // ***** 実行中タスク確認 *****
                // 実行中タスク名取得
                strExecTaskSeqNo = fnc_GetExecTaskSeqNo(STR_STS_EXECUTING);

                // 実行中タスクなければ、順番通り次のタスクを実行
                if (strExecTaskSeqNo == "")
                {
                    // 最後に実行したタスクSeqNo取得
                    strExecTaskSeqNo = fnc_GetExecTaskSeqNo(STR_STS_DONE);
               
                    // 次の実行対象タスクファイル取得
                    strExecTaskSeqNo = fnc_GetNextTaskSeqNo(strExecTaskSeqNo);
                    if (strExecTaskSeqNo == "END")
                    {
                        // 翌日の準備タスクのステータスを更新
                        strTaskExecTimeStamp = DateTime.Now.AddDays(1).ToString("yyyyMMdd");
                        if (fnc_UpdateNextDayTaskStatus(strTaskExecTimeStamp) == true)
                        {
                            strMsg = "準備完了 [<TASK_DAY>], 開始/終了時刻 [<DATE_TIME>]";
                            strMsg = strMsg.Replace("<TASK_DAY>", strTaskExecTimeStamp);
                            strMsg = strMsg.Replace("<DATE_TIME>", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                            CommonLogger.WriteLine(strMsg);
                        }
                    }
                    else if (strExecTaskSeqNo == "NON-START")
                    {
                        // 前日の準備タスクの未完了ステータスを確認
                        strTaskExecTimeStamp = DateTime.Now.AddDays(-1).ToString("yyyyMMdd");
                        strMsg = "準備未完 [<TASK_DAY>], 前日タスク未完 [<DATE_TIME>]";
                        strMsg = strMsg.Replace("<TASK_DAY>", strTaskExecTimeStamp);
                        strMsg = strMsg.Replace("<DATE_TIME>", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                        CommonLogger.WriteLine(strMsg);
                    }
                    else
                    {
                        strExecFile = this.pTaskStauts[strExecTaskSeqNo].task_exe;
                        // タスクファイルを実行
                        if (fnc_ExecBatProcess(strExecFile) == true)
                        {
                            // タスクステータス更新
                            strTaskExecTimeStamp = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            if (fnc_UpdateTaskStatus(strExecTaskSeqNo, STR_STS_EXECUTING, strTaskExecTimeStamp) == true)
                            {
                                // タスクステータス再取得
                                fnc_GetTaskList();

                                strMsg = "タスク実行 [<TASK_SEQ>], 開始時刻 [<DATE_TIME>]";
                                strMsg = strMsg.Replace("<TASK_SEQ>", "SeqNo:(" + strExecTaskSeqNo + ")");
                                strMsg = strMsg.Replace("<DATE_TIME>", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                                CommonLogger.WriteLine(strMsg);
                            }
                        }
                    }                   
                }
                else            
                {
                    // タスクステータス状況確認
                    strTaskExecTimeStamp = fnc_GetTaskEndStatus(strExecTaskSeqNo);
                    if (strTaskExecTimeStamp != "")
                    {
                        fnc_UpdateTaskStatus(strExecTaskSeqNo, STR_STS_DONE, strTaskExecTimeStamp);

                        strMsg = "タスク終了 [<TASK_SEQ>], 終了時刻 [<DATE_TIME>]";
                        strMsg = strMsg.Replace("<TASK_SEQ>", "SeqNo:(" + strExecTaskSeqNo + ")");
                        strMsg = strMsg.Replace("<DATE_TIME>", strTaskExecTimeStamp);

                        CommonLogger.WriteLine(strMsg);
                    }
                    else
                    {
                        strMsg = "タスク実行 [<TASK_SEQ>], タスク実行中";
                        strMsg = strMsg.Replace("<TASK_SEQ>", "SeqNo:(" + strExecTaskSeqNo + ")");
                        CommonLogger.WriteLine(strMsg);
                    }
                }

            }
            catch(Exception ex)
            {
                strMsg = "システムエラー [<SYS_ERR>]";
                strMsg = strMsg.Replace("<SYS_ERR>", ex.Message);
                CommonLogger.WriteLine(strMsg);
            }
            finally
            {
                // 処理終了
                Application.Exit();
            }

        }

        private void frm_main_FormClosing(object sender, FormClosingEventArgs e)
        {
            //// 質問ダイアログを表示する
            //DialogResult result = MessageBox.Show("処理を中止しますか？", "(M-Auto)警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            //if (result == DialogResult.No)
            //{
            //    // はいボタンをクリックしたときはウィンドウを閉じる
            //    e.Cancel = true;
            //}
        }

        private bool fnc_GetTaskList()
        {
            bool blRes;
            OleDbConnection oleConn;
            OleDbCommand oleCmd;
            OleDbDataReader oleReader;
            string strSQL;

            typ_task_status taskStatus;

            try
            {
                oleConn = new OleDbConnection(this.pStrConnecting);
                oleCmd = new OleDbCommand();

                oleConn.Open();
                oleCmd.Connection = oleConn;

                strSQL = "";
                strSQL = strSQL + "SELECT * FROM task_list WHERE 1=1 ";
                strSQL = strSQL + " AND task_day = '" + DateTime.Now.ToString("yyyyMMdd") + "' ";
                strSQL = strSQL + " ORDER BY task_day, task_id";
                oleCmd.CommandText = strSQL;

                this.pTaskStauts = new Dictionary<string, typ_task_status>();
                oleReader = oleCmd.ExecuteReader();
                while (oleReader.Read())
                {
                    taskStatus = new typ_task_status();
                    taskStatus.seq_no = oleReader["seq_no"].ToString();
                    taskStatus.task_id = oleReader["task_id"].ToString();
                    taskStatus.task_day = oleReader["task_day"].ToString();
                    taskStatus.task_status = oleReader["task_status"].ToString();
                    taskStatus.task_start = oleReader["task_start"].ToString();
                    taskStatus.task_end = oleReader["task_end"].ToString();
                    taskStatus.task_status = oleReader["task_status"].ToString();
                    taskStatus.task_exe = oleReader["task_exe"].ToString();
                    taskStatus.task_comment = oleReader["task_comment"].ToString();

                    this.pTaskStauts.Add(taskStatus.seq_no, taskStatus);
                }

                blRes = true;
            }
            catch (Exception ex)
            {
                blRes = false;
                Console.WriteLine(ex.Message);
            }
            return blRes;
        }

        private string fnc_GetNextTaskSeqNo(string strTaskSqeNo)
        {
            string strRes;
            string strNextTaskSeqNo;

            if (strTaskSqeNo == "")
            {
                strRes = "NON-START";
            }
            else
            { 
                // HaspMapにNextタスクが存在しているかどうかをチェック
                strNextTaskSeqNo = (int.Parse(strTaskSqeNo) + 1).ToString("000");

                if (this.pTaskStauts.ContainsKey(strNextTaskSeqNo) == true)
                {
                    strRes = strNextTaskSeqNo;
                }
                else
                {
                    strRes = "END";
                }
            }
            return strRes;
        }

        // 実行中タスクID取得
        private string fnc_GetExecTaskSeqNo(string strStatus)
        {
            string strRes;

            strRes = "";
            try
            {
                foreach (KeyValuePair<string, typ_task_status> item in this.pTaskStauts)
                {
                    if (item.Value.task_status == strStatus)
                    { strRes = item.Key; }
                }
            }
            catch (Exception ex)
            {
                strRes = "";
                Console.WriteLine(ex.Message);
            }
            return strRes;
        }
          
        // BATファイル実行
        private bool fnc_ExecBatProcess(string strBatPath)
        {
            bool blRes;
    
            try
            {
                // ●プロセス起動情報の構築
                ProcessStartInfo startInfo = new ProcessStartInfo();
                // バッチファイルを起動する人は、cmd.exeさんなので
                startInfo.FileName = "cmd.exe";
                // コマンド処理実行後、コマンドウィンドウ終わるようにする。
                //（↓「/c」の後の最後のスペース1文字は重要！）
                startInfo.Arguments = "/c ";
                // コマンド処理であるバッチファイル （ここも最後のスペース重要）
                startInfo.Arguments += strBatPath + " ";
                // ●バッチファイルを別プロセスとして起動
                var proc = Process.Start(startInfo);
                // ●上記バッチ処理が終了するまで待ちます。
                proc.WaitForExit();

                blRes = true;
            }
            catch (Exception ex)
            {
                blRes = false;
                Console.WriteLine(ex.Message);
            }
            return blRes;
        }

        private bool fnc_UpdateTaskStatus(string strTaskSeqNo, string strStatus, string strUpdateTime)
        {
            bool blRes;
            OleDbConnection oleConn;
            OleDbCommand oleCmd;
            string strSQL;

            try
            {
                oleConn = new OleDbConnection(this.pStrConnecting);
                oleCmd = new OleDbCommand();

                oleConn.Open();
                oleCmd.Connection = oleConn;

                strSQL = "";
                strSQL = strSQL + "UPDATE task_list SET ";
                strSQL = strSQL + " task_status ='" + strStatus + "' ";
                // タスク実行開始時刻更新
                if (this.pTaskStauts[strTaskSeqNo].task_status == STR_STS_NOT_START)
                { strSQL = strSQL + ", task_start ='" + strUpdateTime + "' "; }
                // タスク実行終了時刻更新
                if (this.pTaskStauts[strTaskSeqNo].task_status == STR_STS_EXECUTING)
                { strSQL = strSQL + ", task_end ='" + strUpdateTime + "' "; }
                strSQL = strSQL + " WHERE 1=1 ";
                strSQL = strSQL + "  AND seq_no ='" + strTaskSeqNo + "'";

                oleCmd.CommandText = strSQL;
                oleCmd.ExecuteNonQuery();

                blRes = true;
            }
            catch (Exception ex)
            {
                blRes = false;
                Console.WriteLine(ex.Message);
            }
            return blRes;
        }

        private bool fnc_UpdateNextDayTaskStatus(string strNextTaskDay)
        {
            bool blRes;
            OleDbConnection oleConn;
            OleDbCommand oleCmd;
            string strSQL;

            try
            {
                oleConn = new OleDbConnection(this.pStrConnecting);
                oleCmd = new OleDbCommand();

                oleConn.Open();
                oleCmd.Connection = oleConn;

                strSQL = "";
                strSQL = strSQL + "UPDATE task_list SET ";
                strSQL = strSQL + " task_status ='" + STR_STS_DONE + "' ";
                strSQL = strSQL + " , task_start ='" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "' ";
                strSQL = strSQL + " , task_end ='" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "' ";
                strSQL = strSQL + " WHERE 1=1 ";
                strSQL = strSQL + "  AND task_day ='" + strNextTaskDay + "'";
                strSQL = strSQL + "  AND task_id ='000'";

                oleCmd.CommandText = strSQL;
                oleCmd.ExecuteNonQuery();

                blRes = true;
            }
            catch (Exception ex)
            {
                blRes = false;
                Console.WriteLine(ex.Message);
            }
            return blRes;
        }

        // endファイルの更新日付
        private string fnc_GetTaskEndStatus(string strTaskID)
        {
            string strTaskLog;
            string strEndTime;
            bool blRes;
            StreamReader sReader;
            string strTemp;

            strEndTime = "";
            blRes = false;
            try
            {
                // ログファイル取得
                strTaskLog = this.pTaskStauts[strTaskID].task_exe.Replace("bat","log");
                blRes = File.Exists(strTaskLog);

                if (blRes = File.Exists(strTaskLog) == true)
                {
                    // ログファイル読み込み
                    sReader = new StreamReader(strTaskLog);

                    //一行読み込んで表示する
                    while (sReader.Peek() > -1)
                    {
                        strTemp = sReader.ReadLine().ToString().Trim();

                        if (strTemp.Contains("END") == true)
                        {
                            strEndTime = File.GetLastWriteTime(strTaskLog).ToString("yyyy/MM/dd HH:mm:ss");
                            break;
                        }
                    }
                    sReader.Close();                    
                }

                return strEndTime;
            }
            catch (Exception ex)
            {
                strEndTime = "";
                Console.WriteLine(ex.Message);
            }
            return strEndTime;
        }

    }
}
