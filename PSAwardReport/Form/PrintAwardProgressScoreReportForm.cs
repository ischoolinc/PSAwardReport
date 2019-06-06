using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using FISCA.Data;
using Aspose.Cells;
using System.Xml.Linq;
using System.IO;

namespace PSAwardReport.Form
{
    public partial class PrintAwardProgressScoreReportForm : BaseForm
    {
            
        private BackgroundWorker _worker;

        // 目前學年(106、107、108)
        private string school_year;

        // 目前學期(1、2)
        private string semester;
      

        // 使用者 選擇的 系統試別名稱1
        private string _examName1;

        // 使用者 選擇的 系統試別名稱2
        private string _examName2;

        // 使用者 選擇的 系統試別1 ID
        private string _examID1 = "";

        // 使用者 選擇的 系統試別1 ID
        private string _examID2 = "";

           
        public PrintAwardProgressScoreReportForm(List<string> classIDList)
        {
            InitializeComponent();
          
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            buttonX1.Enabled = false; // 關閉按鈕

                 
            // 驗證完畢，開始列印報表
            PrintReport();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void CheckCalculateTermForm_Load(object sender, EventArgs e)
        {

            // 取得系統中所有的設定評量
            List<K12.Data.ExamRecord> examList = K12.Data.Exam.SelectAll();

            foreach (K12.Data.ExamRecord exam in examList)
            {
                comboBoxEx1.Items.Add(exam.Name);

                comboBoxEx2.Items.Add(exam.Name);

                // 建立Exam對照
                if (!ExamDict.ContainsKey(exam.Name))
                {
                    ExamDict.Add(exam.Name, exam.ID);
                }
            }

            comboBoxEx1.SelectedIndex = 0;
            comboBoxEx2.SelectedIndex = 1;


            string courseIDs = string.Join(",", _courseIDList);            
        }


        // 列印 ESL 報表
        private void PrintReport(List<string> courseIDList)
        {
            _examName1 = comboBoxEx1.Text;
            _examName2 = comboBoxEx2.Text;

            _examID1 = ExamDict[_examName1];
            _examID2 = ExamDict[_examName2];

            _worker = new BackgroundWorker();
            _worker.DoWork += new DoWorkEventHandler(Worker_DoWork);
            _worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Worker_RunWorkerCompleted);
            _worker.ProgressChanged += new ProgressChangedEventHandler(Worker_ProgressChanged);
            _worker.WorkerReportsProgress = true;

            _worker.RunWorkerAsync();

        }


        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            _worker.ReportProgress(0, "開始列印 ESL課程進步獎報表...");


            #region 取得目前 系統的 學年度 學期 
            school_year = K12.Data.School.DefaultSchoolYear;
            semester = K12.Data.School.DefaultSemester;
            #endregion
                     

            _worker.ReportProgress(60, "成績排序中...");
          
            _worker.ReportProgress(80, "填寫報表...");

            // 取得 系統預設的樣板
            Workbook wb = new Workbook(new MemoryStream(Properties.Resources.國小進步獎樣板));

            #region 填表

            FillProgressScoreExcelColunm(wb);

            #endregion

            // 把當作樣板的 第一張 移掉
            wb.Worksheets.RemoveAt(0);

            e.Result = wb;

            _worker.ReportProgress(100, "進步獎報表列印完成。");
        }


        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MsgBox.Show("計算進步失敗!!，錯誤訊息:" + e.Error.Message);
                FISCA.Presentation.MotherForm.SetStatusBarMessage(" 進步獎報表產生產生失敗");
                this.Close();
                return;
            }
            else if (e.Cancelled)
            {
                //MsgBox.Show("");
                return;
            }
            else
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage(" 進步獎報表產生完成");
            }



            Workbook wb = (Workbook)e.Result;


            // 電子報表功能先暫時不製做
            #region 電子報表
            //// 檢查是否上傳電子報表
            //if (chkUploadEPaper.Checked)
            //{
            //    List<Document> docList = new List<Document>();
            //    foreach (Section ss in doc.Sections)
            //    {
            //        Document dc = new Document();
            //        dc.Sections.Clear();
            //        dc.Sections.Add(dc.ImportNode(ss, true));
            //        docList.Add(dc);
            //    }

            //    Update_ePaper up = new Update_ePaper(docList, "超額比序項目積分證明單", PrefixStudent.系統編號);
            //    if (up.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
            //    {
            //        MsgBox.Show("電子報表已上傳!!");
            //    }
            //    else
            //    {
            //        MsgBox.Show("已取消!!");
            //    }
            //} 
            #endregion

            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "進步獎報表.xlsx";
            sd.Filter = "Excel 檔案(*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    wb.Save(sd.FileName, SaveFormat.Xlsx);
                    System.Diagnostics.Process.Start(sd.FileName);
                }
                catch
                {
                    MessageBox.Show("檔案儲存失敗");
                }
            }

            this.Close();
        }


        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("" + e.UserState, e.ProgressPercentage);
        }






        

        
        // 填寫 課程成績EXCEL
        private void FillProgressScoreExcelColunm(Workbook wb)
        {
                

            ////把多餘的右半邊CELL欄位 砍掉 (總表)             
            //ws_total.Cells.ClearRange(1, progressScoreCol + 1, totalAwardsCount + 2, 50);
            //ws_total.AutoFitColumns();
            //ws_total.FirstVisibleColumn = 0;// 將打開的介面 調到最左， 要不然就會看到 右邊一片空白。
         

        }

    }



}



