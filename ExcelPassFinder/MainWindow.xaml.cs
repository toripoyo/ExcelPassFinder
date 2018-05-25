using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace xlsFinder
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private bool m_fContinue = false;       // 処理継続フラグ
        private string m_passWd = "";           // 正解パスワード格納用
        private int ProcThreads = 0;            // 並列処理ワーカー数
        private int tmpProcThreads = 0;

        private long procCount = 0;             // 処理したパスワードの数
        private long procDiff = 0;              // 処理速度計算用
        DispatcherTimer timer1;                 // 画面更新用
        private List<Excel.Application> excelApp = new List<Excel.Application>();

        public MainWindow()
        {
            InitializeComponent();
            txtThreads.Text = (Environment.ProcessorCount/2).ToString();       // 初期並列数をコア数/2に設定
            timer1 = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            timer1.Tick += timer1_Tick;
        }

        private async void button_Click(object sender, RoutedEventArgs e)
        {
            if (m_fContinue == false)
            {
                ProcThreads = Convert.ToInt32(txtThreads.Text);
                string fileName = OpenFile();

                // 並列処理ワーカーのインスタンスを作成
                lblOutput.Content = "Creating " + ProcThreads.ToString() + " Workers...";
                tmpProcThreads = ProcThreads;
                for (int worker = 0; worker < tmpProcThreads; worker++)
                {
                    excelApp.Add(new Excel.Application());
                }

                // 処理開始
                m_fContinue = true;
                // 並列でファイルを開きたいので、並列数だけファイルをコピーする
                CopyFiles(fileName, tmpProcThreads);
                System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
                sw.Start();
                timer1.Start();

                // 解読処理開始
                await passUnlockSub("", Convert.ToInt32(txtDigits.Text), fileName);

                sw.Stop();
                timer1.Stop();
                TimeSpan ts = sw.Elapsed;

                // 解析完了
                this.lblOutput.Content = "password is " + m_passWd;
                MessageBox.Show("Completed  " + ts.Hours + ":" + ts.Minutes + ":" + ts.Seconds);
            }
            else
            {
                m_fContinue = false;
                timer1.Stop();
                MessageBox.Show("Aborted");
            }

            // 並列処理ワーカーをクリア
            for (int worker = 0; worker < tmpProcThreads; worker++)
            {
                if (excelApp[worker] != null)
                {
                    excelApp[worker].Quit();
                }
            }
            excelApp.Clear();
        }

        // パスワード解除処理
        private async Task<int> passUnlockSub(string passWd, int pasLen, string fileName)
        {
            // 解析する文字の範囲はまだ指定できません
            char cntChar = (char)0x20;
            if (passWd.Length > pasLen) { return 1; }
            while ((cntChar <= 0x7F) && (m_fContinue == true))
            {
                string tmpPassWd = passWd + Convert.ToString(cntChar);

                if (tmpPassWd.Length != pasLen)
                {
                    if (await passUnlockSub(tmpPassWd, pasLen, fileName) == 0){cntChar++;}
                }
                else
                {
                    await UpdateWorkerStatus(tmpPassWd, fileName);
                    // 次の文字へ
                    cntChar++;
                    tmpPassWd = passWd + Convert.ToString(cntChar);

                    //lblOutput.Content = "Processing... " + tmpPassWd;
                }
            }
            return 0;
        }

        List<Task> OpenBookTask = new List<Task>();
        private async Task<int> UpdateWorkerStatus(string tmpPasswd, string fileName)
        {
            int worker = 0;

            // ワーカーにタスクを登録する
            while (OpenBookTask.Count < ProcThreads)
            {
                OpenBookTask.Add(Task.Run(() => OpenExcelBook(tmpPasswd, fileName + "_" + worker.ToString(), worker)));
                //worker++;
            }
            // タスクを更新する
            while (worker < ProcThreads)
            {
                // そのワーカーのタスクが完了していたら、次のタスクを発行する
                if (OpenBookTask[worker].IsCompleted)
                {
                    OpenBookTask[worker] = Task.Run(() => OpenExcelBook(tmpPasswd, fileName + "_" + worker.ToString(), worker));
                }
                worker++;
            }
            await Task.WhenAny(OpenBookTask);
            return 0;
        }

        private async Task<int> OpenExcelBook(string passWd, string fileName, int workerNum)
        {
            try
            {
                //空白を削除
                passWd = Regex.Replace(passWd, @"\s", "");

                // ワークブックを開く
                // 開けなかったら例外が発生する
                Excel.Workbook excelWork = excelApp[workerNum].Workbooks.Open(
                     fileName, Type.Missing, true, true,
                     passWd, Type.Missing, true,
                     Type.Missing, Type.Missing, Type.Missing,
                     false, Type.Missing, Type.Missing,
                     Type.Missing, Type.Missing);

                // クローズしないほうが速い
                if(excelWork != null) { excelWork.Close(); }               
                excelApp[workerNum].Workbooks.Close();

                // もう一度開く
                excelWork = excelApp[workerNum].Workbooks.Open(
                    fileName, Type.Missing, true, true,
                    passWd, Type.Missing, true,
                    Type.Missing, Type.Missing, Type.Missing,
                    false, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                // 内容を表示する
                excelApp[workerNum].Visible = true;

                // オープン成功したら
                m_fContinue = false;
                m_passWd = passWd;

                return 0;
            }
            catch(Exception e)
            {
                procCount++;
                return -1;
            }
        }

        private string OpenFile()
        {
            //OpenFileDialogクラスのインスタンスを作成
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.FileName = "default.xls";
            ofd.InitialDirectory = @"C:\";
            ofd.Filter = "Excelファイル(*.xls;*.xlsx)|*.xls;*.xlsx|すべてのファイル(*.*)|*.*";
            ofd.FilterIndex = 2;
            ofd.Title = "開くファイルを選択してください";
            ofd.RestoreDirectory = true;
            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;

            //ダイアログを表示する
            bool? result = ofd.ShowDialog();
            if (result == true)
            {
                return ofd.FileName;
            }
            return "";
        }

        // 並列参照用にExcelファイルをコピーする
        // 確か同名ファイルは同時に開けない制約があったと思う
        private void CopyFiles(string fileName, int parNum)
        {
            for (int tmpParNum = 0; tmpParNum < parNum; tmpParNum++)
            {
                System.IO.File.Copy(fileName, fileName + "_" + tmpParNum.ToString(), true);
            }
        }

        // 並列参照用にコピーしたファイルを削除する
        private void DeleteFiles(string fileName, int parNum)
        {
            for (int tmpParNum = 0; tmpParNum < parNum; tmpParNum++)
            {
                System.IO.File.Delete(fileName + "_" + tmpParNum.ToString());
            }
        }

        private void DoEvents()
        {
            DispatcherFrame frame = new DispatcherFrame();
            var callback = new DispatcherOperationCallback(obj =>
            {
                ((DispatcherFrame)obj).Continue = false;
                return null;
            });
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, callback, frame);
            Dispatcher.PushFrame(frame);
        }

        private void updateDisplay()
        {
            lblOutput.Content = "Processing... " + (procCount - procDiff).ToString() + " (p/s)";
            procDiff = procCount;
            DoEvents();
        }

        private void timer1_Tick (object sender, EventArgs e)
        {
            updateDisplay();
        }
    }
}
