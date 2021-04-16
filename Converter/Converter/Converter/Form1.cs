using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;
using System.Collections.Generic;
using System.Xml;
using System.Threading;
using System.Reflection;

namespace Converter
{
    public partial class Form1 : Form
    {
        private Thread ListBoxThread;
        private static string syslogpath = Directory.GetCurrentDirectory() + "\\log";
        private const String filename = "Setting.xml";
        public delegate void AddListItem(String Items);
        public AddListItem myDelegate;
        bool LoadDataFlag = false;
        string folderpath;
        List<string> files = new List<string>();
        bool btnFlag = false;
        private delegate void ListBox(String Items);
        private delegate void ListBoxClear();
        private delegate void ProgessBar(int ToltalCount, int Value);
        private delegate void Label(String msg);
        private delegate void fileProgressBar(int ToltalCount, int Value);
        public Form1()
        {
            InitializeComponent();
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = String.Format("Format Converter Version {0}", version);
            //myDelegate = new AddListItem(ListBoxThreadStart);
        }
        public static bool write_log_timestamp_display(string FileName,string msg, DateTime start, bool executeSpan = false)
        {
            DateTime dispDt = DateTime.Now;
            string DateTimeString = dispDt.ToString(@"MM/dd/yyyy");
            string New_syslogpath = string.Empty;
            string log_file_path = string.Empty;
            string data_str_for_file;
            data_str_for_file = DateTimeString.Replace("/", "_");

            if (New_syslogpath != syslogpath + "\\" + data_str_for_file)
                New_syslogpath = syslogpath + "\\" + data_str_for_file;

            log_file_path = New_syslogpath + "\\" + FileName + ".csv";
            bool Found = false;

            try
            {
                // Check log folder exist
                if (!Directory.Exists(New_syslogpath))
                    Directory.CreateDirectory(New_syslogpath);

                string filter = "*.csv";
                string[] FileMatrix = Directory.GetFileSystemEntries(New_syslogpath, filter);
                string PanelIdString;
                string[] NameMatrix;

                // Check every file name
                for (int i = 0; i < FileMatrix.Length; i++)
                {
                    PanelIdString = Path.GetFileNameWithoutExtension(FileMatrix[i]);
                    NameMatrix = PanelIdString.Split('_');

                    /* Hard coding */
                    //if (NameMatrix[2] == Configuration.ServerRevMsgStruc.PanelId && NameMatrix[1] == Configuration.ServerRevMsgStruc.StationName)
                    //{
                    //    Found = true;
                    //    log_file_path = FileMatrix[i];
                    //    break;
                    //}
                }

                //if (!Found)
                //{
                //    using (FileStream fs = File.Create(log_file_path))
                //    {
                //        fs.Close();
                //    }
                //}

                using (System.IO.StreamWriter file = new System.IO.StreamWriter(log_file_path, true, Encoding.GetEncoding("Big5")))
                {
                    if (!executeSpan)
                    {
                        file.WriteLine(msg);
                    }
                    else
                    {
                        //TimeSpan DTExecutSPAN = start - StationTestItim.Get_time_start();
                        ////UInt32 time = Convert.ToUInt32(DTExecutSPAN.TotalSeconds);
                        //file.WriteLine(DTExecutSPAN.TotalSeconds.ToString("000.00") + "," + msg);
                        //MainFormShowMessage(DTExecutSPAN.TotalSeconds.ToString("000.00") + "," + msg, HIGH);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                //DisplayException(ex);
                return false;
            }
        }
        private static string[] GetFileNames(string path, string filter)
        {
            string[] files = Directory.GetFiles(path, filter);
            for (int i = 0; i < files.Length; i++)
                files[i] = Path.GetFileName(files[i]);
            return files;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ListBoxThread = new Thread(new ThreadStart(ConvertFunc));
            ListBoxThread.Start();
        }
        public void ConvertFunc()
        {
            try
            {
                string FileStr;
                string NewString = "";
                LabelMsg("Start!");
                XmlDocument doc = new XmlDocument();
                doc.Load(filename);
                XmlNode Search_node = doc.DocumentElement.SelectSingleNode("/Root/Search");
                XmlNode Arrange_node = doc.DocumentElement.SelectSingleNode("/Root/Arrangement");
                ListBoxThreadClear();
                LabelMsg("Processing...");
                
                for (int ix = 0; ix < files.Count(); ix++)
                {
                    FileStr = folderpath + "\\" + files[ix];
                    string[] Arr_node = Arrange_node.InnerText.ToString().Split(',');
                    var reader = new StreamReader(File.OpenRead(FileStr));
                    var line = "";
                    int count = 0;
                    int finishedline = 0;
                    while (!reader.EndOfStream)
                    {
                        count = System.IO.File.ReadAllLines(FileStr).Length;
                        line = reader.ReadLine();
                        string[] values = line.Split(',');
                        if (values[2] == Search_node.InnerText)
                        {
                            NewString = "";
                            for (int i = 0; i < Arr_node.Length; i++)
                            {
                                NewString = NewString + values[System.Convert.ToInt16(Arr_node[i])] + ",";
                            }
                            //NewString = values[0] + "," + values[4] + "," + values[3] + "," + values[5] + "," + values[6] + "," + values[7] + "," + values[8] + "," + values[9] + "," + values[10] + "," + values[11] + "," + values[12] + "," + values[13] + "," + values[14] + "," + values[15] + "," + values[16];
                            write_log_timestamp_display(files[ix], NewString, DateTime.Now);
                            NewString = "";
                        }
                        else
                        {
                            NewString = line;
                            write_log_timestamp_display(files[ix], NewString, DateTime.Now);
                        }
                        finishedline++;
                        DofileProgress(count, finishedline);
                    }
                    DoProgress(files.Count(), ix);
                    ListBoxThreadStart(files[ix] + "  ,Finished");
                }
                LabelMsg("Completed!");
            }              
            catch (IOException e)
            {
                Console.WriteLine($"The file could not be opened: '{e}'");
            }
        }
        private void LabelMsg(String msg)
        {
            //listBox1.Items.Add(Items);
            if (this.InvokeRequired) // 若非同執行緒
            {
                Label Lab = new Label(LabelMsg); //利用委派執行
                this.Invoke(Lab, msg);
            }
            else // 同執行緒
            {
                this.label1.Text = msg;
                //this.textBox1.Text += sMessage + Environment.NewLine;
            }
        }
        private void startProgress(int ToltalCount, int Value)
        {
            // 顯示進度條控制元件.
            progressBar1.Visible = true;
            // 設定進度條最小值.
            progressBar1.Minimum = 1;
            // 設定進度條最大值.
            progressBar1.Maximum = ToltalCount+1;
            // 設定進度條初始值
            progressBar1.Value = Value + 1;
            // 設定每次增加的步長
            progressBar1.Step = 1;
            // 迴圈執行
            progressBar1.PerformStep();
            //progressBar1.Visible = false;
        }
        private void DoProgress(int ToltalCount, int Value)
        {
            if (this.InvokeRequired)
            {
                ProgessBar PBar = new ProgessBar(DoProgress);
                this.Invoke(PBar, ToltalCount, Value);
            }
            else
            {
                this.startProgress(ToltalCount, Value);
            }
        }
        private void fileProgress(int ToltalCount, int Value)
        {
            // 顯示進度條控制元件.
            progressBar2.Visible = true;
            // 設定進度條最小值.
            progressBar2.Minimum = 1;
            // 設定進度條最大值.
            progressBar2.Maximum = ToltalCount + 1;
            // 設定進度條初始值
            progressBar2.Value = Value + 1;
            // 設定每次增加的步長
            progressBar2.Step = 1;
            // 迴圈執行
            progressBar2.PerformStep();
            //progressBar1.Visible = false;
        }
        private void DofileProgress(int ToltalCount, int Value)
        {
            if (this.InvokeRequired)
            {
                fileProgressBar PBar = new fileProgressBar(DofileProgress);
                this.Invoke(PBar, ToltalCount, Value);
            }
            else
            {
                this.fileProgress(ToltalCount, Value);
            }
        }
        private void ListBoxThreadStart(String Items)
        {
            //listBox1.Items.Add(Items);
            if (this.InvokeRequired) // 若非同執行緒
            {
                ListBox LBox = new ListBox(ListBoxThreadStart); //利用委派執行
                this.Invoke(LBox, Items);
            }
            else // 同執行緒
            {
                this.listBox1.Items.Add(Items);
                //this.textBox1.Text += sMessage + Environment.NewLine;
            }
        }
        private void ListBoxThreadClear()
        {
            //listBox1.Items.Add(Items);
            if (this.InvokeRequired) // 若非同執行緒
            {
                ListBoxClear LBoxClear = new ListBoxClear(ListBoxThreadClear); //利用委派執行
                this.Invoke(LBoxClear);
            }
            else // 同執行緒
            {
                this.listBox1.Items.Clear();
                //this.textBox1.Text += sMessage + Environment.NewLine;
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FolderBrowserDialog = new FolderBrowserDialog();
            LoadDataFlag = false;
            DialogResult result = FolderBrowserDialog.ShowDialog();
            folderpath = FolderBrowserDialog.SelectedPath;
            string[] csvFiles = GetFileNames(folderpath, "*.csv");
            foreach (string csvFile in csvFiles)
            {
                if (csvFile.Contains("CBMS"))
                {
                    files.Add(csvFile);
                    listBox2.Items.Add(csvFile);
                }
            }
        }
    }
}
