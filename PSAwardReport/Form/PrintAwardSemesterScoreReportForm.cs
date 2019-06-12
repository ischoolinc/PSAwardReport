using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Xml;
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


    public partial class PrintAwardSemesterScoreReportForm : BaseForm
    {

        private BackgroundWorker _worker;

        // 目前學年(106、107、108)
        private string _school_year;

        // 目前學期(1、2)
        private string _semester;
      
        // 儲放 所有課程的科目名稱
        private List<string> _subjectList = new List<string>();

        // 儲放 所有選擇課程的科目名稱
        private List<string> _selsubjectList = new List<string>();

        // 使用者所選的班級清單
        private List<K12.Data.ClassRecord> _classList;

        // 學生清單
        private List<K12.Data.StudentRecord> _studentList;

        // 學生Dict (classID,List<StudentRecord>)
        private Dictionary<string, List<K12.Data.StudentRecord>> _studentDict;

        // 課程清單(對照出 節數/權數使用)
        private List<K12.Data.CourseRecord> _courseList;

        // 學期科目成績清單
        private List<K12.Data.SemesterScoreRecord> _subjectScoreList;

        // 評量成績 (studentID,List<SubjectScore>)
        private Dictionary<string, List<K12.Data.SubjectScore>> _subjectScoreDict;

        // 總成績比較 (studentID,List<AwardScore>)
        private Dictionary<string, List<AwardScore>> _awardScoreDict;

        // 使用者所選的科目數量
        private int selectSubjectCount = 0;

        public PrintAwardSemesterScoreReportForm(List<K12.Data.ClassRecord> classList)
        {

            _studentDict = new Dictionary<string, List<K12.Data.StudentRecord>>();

            _subjectScoreDict = new Dictionary<string, List<K12.Data.SubjectScore>>();

            _awardScoreDict = new Dictionary<string, List<AwardScore>>();

            InitializeComponent();

            // 取得預設學年度學期
            _school_year = K12.Data.School.DefaultSchoolYear;

            _semester = K12.Data.School.DefaultSemester;

            _classList = classList;

        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            buttonX1.Enabled = false; // 關閉按鈕

            buttonX2.Enabled = false;

            // 使用者是否有選擇科目
            bool hasSubject = false;

            selectSubjectCount = 0;

            foreach (ListViewItem lvi in lvSubject.Items)
            {
                if (lvi.Checked)
                {
                    // 加入已選擇科目清單
                    _selsubjectList.Add(lvi.Text);

                    // 使用者選擇科目數量
                    selectSubjectCount++;
                    hasSubject = true;
                }
            }

            if (!hasSubject)
            {
                MsgBox.Show("必須選擇科目！");
                return;
            };

            // 驗證完畢，開始列印報表
            PrintReport();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            // 取消原本的工作
            _worker.CancelAsync();

            this.Close();
        }


        private void PrintReport()
        {

            _worker = new BackgroundWorker();
            _worker.DoWork += new DoWorkEventHandler(Worker_DoWork);
            _worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Worker_RunWorkerCompleted);
            _worker.ProgressChanged += new ProgressChangedEventHandler(Worker_ProgressChanged);
            _worker.WorkerReportsProgress = true;
            _worker.WorkerSupportsCancellation = true; // 要設定這一個 ， 才可以取消中止

            _worker.RunWorkerAsync();

        }


        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            _worker.ReportProgress(0, "開始列印 五育獎報表...");

            _worker.ReportProgress(20, "取得本學期開設的班級課程...");

            // 本學期 班級開設的 課程清單
            _courseList = K12.Data.Course.SelectByClass(int.Parse(_school_year), int.Parse(_semester), _classList);

            // 課程ID List
            List<string> courseIDList = new List<string>();

            foreach (K12.Data.CourseRecord cr in _courseList)
            {
                courseIDList.Add(cr.ID);
            }



            _worker.ReportProgress(40, "取得本學期科目成績...");

            // 取得學生清單(依班級)
            _studentList = K12.Data.Student.SelectByClasses(_classList);

            // 學生ID List
            List<string> stuIDList = new List<string>();

            //整理學生清單(依照班級ID 分類)
            foreach (K12.Data.StudentRecord sr in _studentList)
            {
                // 若非一般生 不列入排名學生清單
                if (sr.Status != K12.Data.StudentRecord.StudentStatus.一般)
                {
                    continue;
                }

                if (!_studentDict.ContainsKey(sr.Class.ID))
                {
                    _studentDict.Add(sr.Class.ID, new List<K12.Data.StudentRecord>());

                    _studentDict[sr.Class.ID].Add(sr);
                }
                else
                {
                    _studentDict[sr.Class.ID].Add(sr);

                }

                stuIDList.Add(sr.ID);
            }

            // 排序(依照座號)

            foreach (string classID in _studentDict.Keys)
            {

                _studentDict[classID].Sort((x, y) => { return x.SeatNo.HasValue ? x.SeatNo.Value.CompareTo(y.SeatNo) : -1; });

            }

            // 取得 這群學生本學期 評量成績
            _subjectScoreList = K12.Data.SemesterScore.SelectBySchoolYearAndSemester(stuIDList, int.Parse(_school_year), int.Parse(_semester));

            // 整理學期科目成績
            foreach (K12.Data.SemesterScoreRecord semesterRecord in _subjectScoreList)
            {
                if (semesterRecord.Subjects.Keys.Count > 0)
                {
                    foreach (string subject in semesterRecord.Subjects.Keys)
                    {
                        // 只有在本次選擇的科目才納入
                        if (!_selsubjectList.Contains(subject))
                        {
                            continue;

                        }

                        if (!_subjectScoreDict.ContainsKey(semesterRecord.RefStudentID))
                        {
                            _subjectScoreDict.Add(semesterRecord.RefStudentID, new List<K12.Data.SubjectScore>());

                            _subjectScoreDict[semesterRecord.RefStudentID].Add(semesterRecord.Subjects[subject]);
                        }
                        else
                        {
                            _subjectScoreDict[semesterRecord.RefStudentID].Add(semesterRecord.Subjects[subject]);
                        }
                    }
                }
            }


            _worker.ReportProgress(60, "成績排序中...");


            int progress = 100 / (_subjectScoreDict.Keys.Count != 0 ? _subjectScoreDict.Keys.Count : 100);
            int studentCount = 0;

            foreach (string stuID in _subjectScoreDict.Keys)
            {
                if (!_awardScoreDict.ContainsKey(stuID))
                {
                    _awardScoreDict.Add(stuID, new List<AwardScore>());

                    foreach (K12.Data.SubjectScore subjectRecord in _subjectScoreDict[stuID])
                    {
                        AwardScore awardScore;

                        awardScore = _awardScoreDict[stuID].Find(awardscore => awardscore.ScoreType == "學期總成績");

                        // 舊的沒有的話， 第一筆 新增
                        if (awardScore == null)
                        {
                            awardScore = new AwardScore();

                            decimal? credit = subjectRecord.Credit;

                            awardScore.ScoreType = "學期總成績";

                            awardScore.Score = subjectRecord.Score * credit;

                            _awardScoreDict[stuID].Add(awardScore);
                        }
                        // 已經有成績的話 繼續 加總
                        else
                        {
                            decimal? credit = subjectRecord.Credit;
                            awardScore.Score += subjectRecord.Score * credit; 
                           
                        }
                    }
                }

                studentCount++;

                _worker.ReportProgress(progress * studentCount, "計算總成績中...");
            }




            _worker.ReportProgress(70, "填寫報表...");

            // 取得 系統預設的樣板
            Workbook wb = new Workbook(new MemoryStream(Properties.Resources.國小五育獎樣板));

            #region 填表

            FillProgressScoreExcelColunm(wb);

            #endregion

            //// 把當作樣板的 第一張 移掉
            wb.Worksheets.RemoveAt("樣板");

            wb.Worksheets.RemoveAt("理想產生結果");

            e.Result = wb;

            _worker.ReportProgress(100, "五育獎報表列印完成。");
        }


        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MsgBox.Show("計算進步失敗!!，錯誤訊息:" + e.Error.Message);
                FISCA.Presentation.MotherForm.SetStatusBarMessage(" 五育獎報表產生產生失敗");
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
                FISCA.Presentation.MotherForm.SetStatusBarMessage(" 五育獎報表產生完成");
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
            sd.FileName = "五育獎報表.xlsx";
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



        // 填 EXCEL
        private void FillProgressScoreExcelColunm(Workbook wb)
        {
            // 一個班級　 開一個 Worksheet
            foreach (K12.Data.ClassRecord classRecord in _classList)
            {
                #region 填寫表頭
                Worksheet ws = wb.Worksheets[wb.Worksheets.Add()];

                // 複製樣板
                ws.Copy(wb.Worksheets["樣板"]);

                ws.Name = classRecord.Name + "_五育獎";

                // 填表頭 (座號 姓名 兩欄 、兩次評量各自的科目數  最右邊最後比較結果 二欄)
                Range headRange = ws.Cells.CreateRange(0, 0, 2, 2 + selectSubjectCount + 2);

                headRange.SetStyle(wb.Worksheets["樣板"].Cells["A1"].GetStyle());

                // 右邊總成績
                Range resultHeadRange = ws.Cells.CreateRange(2, 2 + selectSubjectCount, 2, 2);

                resultHeadRange.Copy(wb.Worksheets["樣板"].Cells.CreateRange("D3", "E4"));

                resultHeadRange.ColumnWidth = 10;

                #endregion

                // 填科目 、填分數類別
                int subjectPlace = 0;

                foreach (string subject in _selsubjectList)
                {
                    // 科目名稱全名，後面帶權數 EX: 國文(3) 
                    string subjectFullName = "";

                    // 這個班級的課程
                    subjectFullName = _courseList.Find(c => c.Class.ID == classRecord.ID && c.Subject == subject) != null ? subject + "(" + _courseList.Find(c => c.Class.ID == classRecord.ID && c.Subject == subject).Credit + ")" : subject + "(?)";

                    // 取得該課程的評分樣版 查詢 定期評量、平時評量的比例
                    K12.Data.AssessmentSetupRecord assessmentsetup = _courseList.Find(c => c.Class.ID == classRecord.ID && c.Subject == subject) != null ? _courseList.Find(c => c.Class.ID == classRecord.ID && c.Subject == subject).AssessmentSetup : null;

                    //科目
                    Cell cell_subect = ws.Cells[2, 2 + subjectPlace];

                    cell_subect.Copy(wb.Worksheets["樣板"].Cells["C3"]);

                    cell_subect.Value = subjectFullName;

                    // 填 學期成績 字樣
                     ws.Cells[3, 2 + subjectPlace].Value ="學期成績";

                    ws.Cells[3, 2 + subjectPlace].Copy(wb.Worksheets["樣板"].Cells["C4"]);


                    subjectPlace++; 
                }


                // 整理的學生清單 沒有 班級， 代表本班級沒有學生 跳過
                if (!_studentDict.ContainsKey(classRecord.ID))
                {
                    continue;
                }

                int progress = 100 / (_studentDict[classRecord.ID].Count != 0 ? _studentDict[classRecord.ID].Count : 100);
                int studentCount = 0;
                //填學生資料
                foreach (K12.Data.StudentRecord sr in _studentDict[classRecord.ID])
                {

                    // 座號
                    ws.Cells[4 + studentCount, 0].Value = sr.SeatNo;

                    // 姓名
                    ws.Cells[4 + studentCount, 1].Value = sr.Name;


                    // 填表格樣式(黃)
                    Range dataRange = ws.Cells.CreateRange(4 + studentCount, 0, 1, 2 + selectSubjectCount);

                    dataRange.SetStyle(wb.Worksheets["樣板"].Cells["A5"].GetStyle());

                    // 填表格樣式(右邊 比較結果)
                    Range resultRange = ws.Cells.CreateRange(4 + studentCount, 2 + selectSubjectCount , 1, 2);

                    resultRange.Copy(wb.Worksheets["樣板"].Cells.CreateRange("D5", "E5"));


                    // 填 學期科目成績
                    if (_subjectScoreDict.ContainsKey(sr.ID))
                    {
                        foreach (K12.Data.SubjectScore subjectScore in _subjectScoreDict[sr.ID])
                        {
                            // 科目全名 (EX: 國文(3))
                            string subjectFullName = subjectScore.Subject + "(" + subjectScore.Credit + ")";

                            for (int i = 2; i <= selectSubjectCount +1; i++)
                            {
                                if ("" + ws.Cells[2, i].Value == subjectFullName)
                                {
                                    // 學期科目成績 
                                    ws.Cells[4 + studentCount, i].Value = subjectScore.Score;                                    
                                }
                            }
                        }


                    }

                    // 填總成績 (總成績、排名)
                    if (_awardScoreDict.ContainsKey(sr.ID))
                    {
                        foreach (AwardScore awardScore in _awardScoreDict[sr.ID])
                        {
                            // 學期總成績
                            if (awardScore.ScoreType == "" + ws.Cells[3, 2 + selectSubjectCount].Value)
                            {
                                // 學期總成績 
                                ws.Cells[4 + studentCount, 2 + selectSubjectCount].Value = awardScore.Score;

                                // 學期總成績 第一位同學 cell Name
                                string firstCellName = ws.Cells[4, 2 + selectSubjectCount].Name;

                                // 學期總成績 最後一位同學 cell Name
                                string lastCellName = ws.Cells[4 + _studentDict[classRecord.ID].Count - 1, 2 + selectSubjectCount].Name;

                                // 學期總成績 目前此為同學 cell Name
                                string nowCellName = ws.Cells[4 + studentCount, 2 + selectSubjectCount].Name;

                                // 排名(用EXCEL 公式算，同分 同名次 不接續排名)
                                ws.Cells[4 + studentCount, 2 + selectSubjectCount + 1].Formula = "=RANK(" + nowCellName + "," + firstCellName + ":" + lastCellName + ")";
                            }

                        }
                    }

                    studentCount++;

                    _worker.ReportProgress(progress * studentCount, "填寫" + classRecord.Name + "五育獎報表");
                }

                ws.FirstVisibleColumn = 0;// 將打開的介面 調到最左， 要不然就會看到 右邊一片空白。
            }




        }

        // 全選科目
        private void chkSubjSelAll_CheckedChanged(object sender, EventArgs e)
        {
            foreach (ListViewItem lvi in lvSubject.Items)
            {
                lvi.Checked = chkSubjSelAll.Checked;
            }
        }


        private void LoadSubject()
        {
            lvSubject.Items.Clear();
            string ExamID = "";

            //foreach (string examName  in ExamDict.Keys)
            //{
            //    if (examName == cboExam.Text)
            //    {
            //        ExamID = ExamDict[examName];
            //        break;
            //    }
            //}

            foreach (string subjName in _subjectList)
            {
                lvSubject.Items.Add(subjName);
            }

        }

        private void SettingForm_Load(object sender, EventArgs e)
        {
            QueryHelper qh = new QueryHelper();
            #region 抓科目

            // 選出 本學期課程 的科目名稱         
            string sql = @"
    SELECT 		
	DISTINCT	course.subject
	FROM course 
	WHERE 	
    course.school_year = " + _school_year + @"
    AND course.semester = " + _semester + @"
	AND course.subject IS NOT NULL
    ORDER BY course.subject DESC";

            DataTable dt = qh.Select(sql);

            foreach (DataRow row in dt.Rows)
            {
                string subject = "" + row["subject"];
                _subjectList.Add(subject);
            }




            LoadSubject();

            #endregion

            circularProgress1.Visible = false;
        }


    }



}




