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
    public partial class PrintAwardProgressScoreReportForm : BaseForm
    {

        private BackgroundWorker _worker;

        // 目前學年(106、107、108)
        private string _school_year;

        // 目前學期(1、2)
        private string _semester;


        // 使用者 選擇的 系統試別名稱1
        private string _examName1;

        // 使用者 選擇的 系統試別名稱2
        private string _examName2;

        // 使用者 選擇的 系統試別1 ID
        private string _examID1 = "";

        // 使用者 選擇的 系統試別1 ID
        private string _examID2 = "";

        // 系統Exam 中文名稱 對應 Exam ID
        Dictionary<string, string> ExamDict = new Dictionary<string, string>();

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

        // 評量成績清單
        private List<K12.Data.SCETakeRecord> _sceList;

        // 評量成績 (studentID,List<SCETakeRecord>)
        private Dictionary<string, List<K12.Data.SCETakeRecord>> _sceDict;

        // 總成績比較 (studentID,List<AwardScore>)
        private Dictionary<string, List<AwardScore>> _awardScoreDict;

        // 使用者所選的科目數量
        private int selectSubjectCount = 0;

        public PrintAwardProgressScoreReportForm(List<K12.Data.ClassRecord> classList)
        {

            _studentDict = new Dictionary<string, List<K12.Data.StudentRecord>>();

            _sceDict = new Dictionary<string, List<K12.Data.SCETakeRecord>>();

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
            this.Close();
        }



        // 列印 ESL 報表
        private void PrintReport()
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

            _worker.ReportProgress(20, "取得本學期開設的班級課程...");

            // 本學期 班級開設的 課程清單
            _courseList = K12.Data.Course.SelectByClass(int.Parse(_school_year), int.Parse(_semester), _classList);

            // 課程ID List
            List<string> courseIDList = new List<string>();

            foreach (K12.Data.CourseRecord cr in _courseList)
            {
                courseIDList.Add(cr.ID);
            }



            _worker.ReportProgress(40, "取得本學期評量成績...");

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

            // 取得評量成績
            _sceList = K12.Data.SCETake.SelectByStudentAndCourse(stuIDList, courseIDList);

            // 整理評量成績
            foreach (K12.Data.SCETakeRecord scetRecord in _sceList)
            {
                // 只有在本次選擇的科目才納入
                if (!_selsubjectList.Contains(scetRecord.Course.Subject))
                {
                    continue;
                }

                if (!_sceDict.ContainsKey(scetRecord.RefStudentID))
                {
                    _sceDict.Add(scetRecord.RefStudentID, new List<K12.Data.SCETakeRecord>());

                    _sceDict[scetRecord.RefStudentID].Add(scetRecord);
                }
                else
                {
                    _sceDict[scetRecord.RefStudentID].Add(scetRecord);
                }
            }


            _worker.ReportProgress(60, "成績排序中...");


            foreach (string stuID in _sceDict.Keys)
            {
                if (!_awardScoreDict.ContainsKey(stuID))
                {
                    _awardScoreDict.Add(stuID, new List<AwardScore>());

                    foreach (K12.Data.SCETakeRecord scetRecord in _sceDict[stuID])
                    {
                        // 第一次評量成績
                        if (scetRecord.Exam.Name == _examName1)
                        {                            
                            AwardScore awardScore = new AwardScore();

                            decimal? credit = scetRecord.Course.Credit;

                            K12.Data.AssessmentSetupRecord assessmentsetup = scetRecord.Course.AssessmentSetup;

                            // 定期的比例
                            int scoreRatio = int.Parse(GetScoreRatio(assessmentsetup));

                            // 平時的比例 (兩者相加 為100)
                            int assignmentScoreRatio = 100 - scoreRatio;

                            awardScore.ScoreType = _examName1 + "總成績";

                            if (awardScore.Score.HasValue)
                            {
                                awardScore.Score += (int.Parse(GetScore(scetRecord)) * scoreRatio + int.Parse(GetAssignmentScore(scetRecord)) * assignmentScoreRatio) * credit;
                            }
                            else
                            {
                                awardScore.Score = (int.Parse(GetScore(scetRecord)) * scoreRatio + int.Parse(GetAssignmentScore(scetRecord)) * assignmentScoreRatio) * credit;
                            }
                            
                            _awardScoreDict[stuID].Add(awardScore);
                        }

                        // 第二次評量成績
                        if (scetRecord.Exam.Name == _examName2)
                        {
                            AwardScore awardScore = new AwardScore();

                            decimal? credit = scetRecord.Course.Credit;

                            K12.Data.AssessmentSetupRecord assessmentsetup = scetRecord.Course.AssessmentSetup;

                            // 定期的比例
                            int scoreRatio = int.Parse(GetScoreRatio(assessmentsetup));

                            // 平時的比例 (兩者相加 為100)
                            int assignmentScoreRatio = 100 - scoreRatio;

                            awardScore.ScoreType = _examName2 + "總成績";

                            if (awardScore.Score.HasValue)
                            {
                                awardScore.Score += (int.Parse(GetScore(scetRecord)) * scoreRatio + int.Parse(GetAssignmentScore(scetRecord)) * assignmentScoreRatio) * credit;
                            }
                            else
                            {
                                awardScore.Score = (int.Parse(GetScore(scetRecord)) * scoreRatio + int.Parse(GetAssignmentScore(scetRecord)) * assignmentScoreRatio) * credit;
                            }

                            _awardScoreDict[stuID].Add(awardScore);

                        }
                    }
                }

            }




            _worker.ReportProgress(70, "填寫報表...");

            // 取得 系統預設的樣板
            Workbook wb = new Workbook(new MemoryStream(Properties.Resources.國小進步獎樣板));

            #region 填表

            FillProgressScoreExcelColunm(wb);

            #endregion

            //// 把當作樣板的 第一張 移掉
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

                ws.Name = classRecord.Name + "_進步獎";

                // 填表頭 (座號 姓名 兩欄 、兩次評量各自的科目數 乘四欄 最右邊最後比較結果 六欄)
                Range headRange = ws.Cells.CreateRange(0, 0, 2, 2 + selectSubjectCount * 4 + 6);

                headRange.SetStyle(wb.Worksheets["樣板"].Cells["A1"].GetStyle());

                // 填評量名稱

                // 評量1
                Cell cell_exam1 = ws.Cells[2, 2];

                cell_exam1.Copy(wb.Worksheets["樣板"].Cells["C3"]);

                cell_exam1.Value = _examName1;

                // 要被併的那些格 要另外設定 Styele 要不然會出現原始預設的格子(EX: 沒有框線)
                Range examRange1 = ws.Cells.CreateRange(2, 2, 1, selectSubjectCount * 2);

                examRange1.SetStyle(wb.Worksheets["樣板"].Cells["C3"].GetStyle());


                // 每有選一科　就會占兩格（定期評量、平時評量）
                ws.Cells.Merge(2, 2, 1, selectSubjectCount * 2);


                // 評量2
                Cell cell_exam2 = ws.Cells[2, 2 + selectSubjectCount * 2];

                cell_exam2.Copy(wb.Worksheets["樣板"].Cells["E3"]);

                cell_exam2.Value = _examName2;

                // 要被併的那些格 要另外設定 Styele 要不然會出現原始預設的格子(EX: 沒有框線)              
                Range examRange2 = ws.Cells.CreateRange(2, 2 + selectSubjectCount * 2, 1, selectSubjectCount * 2);

                examRange2.SetStyle(wb.Worksheets["樣板"].Cells["E3"].GetStyle());


                // 每有選一科　就會占兩格（定期評量、平時評量）
                ws.Cells.Merge(2, 2 + selectSubjectCount * 2, 1, selectSubjectCount * 2);


                // 右邊總成績
                Range resultHeadRange = ws.Cells.CreateRange(2, 2 + selectSubjectCount * 4, 3, 6);

                resultHeadRange.Copy(wb.Worksheets["樣板"].Cells.CreateRange("G3", "L5"));

                // 依照選擇兩次試別 更改顯示文字
                ws.Cells[4, 2 + selectSubjectCount * 4].Value =_examName1 +"總成績";

                ws.Cells[4, 2 + selectSubjectCount * 4 +2].Value = _examName2 + "總成績";

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

                    // 定期的比例
                    int scoreRatio = int.Parse(GetScoreRatio(assessmentsetup));

                    // 平時的比例 (兩者相加 為100)
                    int assignmentScoreRatio = 100 - scoreRatio;

                    // 評量1　科目
                    Cell cell_exam_subect1 = ws.Cells[3, 2 + subjectPlace];

                    cell_exam_subect1.Copy(wb.Worksheets["樣板"].Cells["C3"]);

                    cell_exam_subect1.Value = subjectFullName;

                    // 要被併的那一格 要另外設定 Styele 要不然會出現原始預設的格子(EX: 沒有框線)
                    ws.Cells[3, 2 + subjectPlace].SetStyle(wb.Worksheets["樣板"].Cells["C3"].GetStyle());

                    // 科目　占兩格
                    ws.Cells.Merge(3, 2 + subjectPlace, 1, 2);

                    // 評量1 分數類別1 (定期評量)
                    Cell cell_exam1_score1 = ws.Cells[4, 2 + subjectPlace];

                    cell_exam1_score1.Copy(wb.Worksheets["樣板"].Cells["C5"]);

                    cell_exam1_score1.Value = cell_exam1_score1.Value + "(" + scoreRatio + "%)";

                    // 評量1 分數類別2　(平時評量)
                    Cell cell_exam1_score2 = ws.Cells[4, 2 + subjectPlace + 1];

                    cell_exam1_score2.Copy(wb.Worksheets["樣板"].Cells["D5"]);

                    cell_exam1_score2.Value = cell_exam1_score2.Value + "(" + assignmentScoreRatio + "%)";


                    // 評量2　科目
                    Cell cell_exam_subect2 = ws.Cells[3, 2 + subjectPlace + selectSubjectCount * 2];

                    cell_exam_subect2.Copy(wb.Worksheets["樣板"].Cells["E3"]);

                    cell_exam_subect2.Value = subjectFullName;

                    // 要被併的那一格 要另外設定 Styele 要不然會出現原始預設的格子(EX: 沒有框線)
                    ws.Cells[3, 2 + subjectPlace + selectSubjectCount * 2 + 1].SetStyle(wb.Worksheets["樣板"].Cells["E3"].GetStyle());

                    // 科目　占兩格
                    ws.Cells.Merge(3, 2 + subjectPlace + selectSubjectCount * 2, 1, 2);

                    // 評量2 分數類別1 (定期評量)
                    Cell cell_exam2_score1 = ws.Cells[4, 2 + subjectPlace + selectSubjectCount * 2];

                    cell_exam2_score1.Copy(wb.Worksheets["樣板"].Cells["E5"]);

                    cell_exam2_score1.Value = cell_exam2_score1.Value + "(" + scoreRatio + "%)";

                    // 評量2 分數類別2　(平時評量)
                    Cell cell_exam2_score2 = ws.Cells[4, 2 + subjectPlace + selectSubjectCount * 2 + 1];

                    cell_exam2_score2.Copy(wb.Worksheets["樣板"].Cells["F5"]);

                    cell_exam2_score2.Value = cell_exam2_score2.Value + "(" + assignmentScoreRatio + "%)";

                    subjectPlace = subjectPlace + 2;
                }


                // 整理的學生清單 沒有 班級， 代表本班級沒有學生 跳過
                if (!_studentDict.ContainsKey(classRecord.ID))
                {
                    continue;
                }

                int progress = 100 / _studentDict[classRecord.ID].Count;
                int studentCount = 0;
                //填學生資料
                foreach (K12.Data.StudentRecord sr in _studentDict[classRecord.ID])
                {

                    // 座號
                    ws.Cells[5 + studentCount, 0].Value = sr.SeatNo;

                    // 姓名
                    ws.Cells[5 + studentCount, 1].Value = sr.Name;


                    // 填表格樣式(黃)
                    Range dataRange = ws.Cells.CreateRange(5 + studentCount, 0, 1, 2 + selectSubjectCount * 4);

                    dataRange.SetStyle(wb.Worksheets["樣板"].Cells["A6"].GetStyle());

                    // 填表格樣式(右邊 比較結果)
                    Range resultRange = ws.Cells.CreateRange(5 + studentCount, 2 + selectSubjectCount * 4, 1, 6);

                    resultRange.Copy(wb.Worksheets["樣板"].Cells.CreateRange("G6", "L6"));

                    // 填 評量成績
                    if (_sceDict.ContainsKey(sr.ID))
                    {
                        foreach (K12.Data.SCETakeRecord scetRecord in _sceDict[sr.ID])
                        {
                            // 此 成績 屬於 評量成績1
                            if (scetRecord.Exam.Name == _examName1)
                            {
                                // 科目全名 (EX: 國文(3))
                                string subjectFullName = scetRecord.Course.Subject + "(" + scetRecord.Course.Credit + ")";


                                for (int i = 2; i <= selectSubjectCount * 2; i = i + 2)
                                {
                                    if ("" + ws.Cells[3, i].Value == subjectFullName)
                                    {
                                        // 定期評量
                                        ws.Cells[5 + studentCount, i].Value = GetScore(scetRecord);
                                        // 平時評量
                                        ws.Cells[5 + studentCount, i + 1].Value = GetAssignmentScore(scetRecord);
                                    }
                                }
                            }
                            // 此 成績 屬於 評量成績2
                            else if (scetRecord.Exam.Name == _examName2)
                            {
                                // 科目全名 (EX: 國文(3))
                                string subjectFullName = scetRecord.Course.Subject + "(" + scetRecord.Course.Credit + ")";

                                for (int i = 2; i <= selectSubjectCount * 2; i = i + 2)
                                {
                                    if ("" + ws.Cells[3, i + selectSubjectCount * 2].Value == subjectFullName)
                                    {
                                        // 定期評量
                                        ws.Cells[5 + studentCount, i + selectSubjectCount * 2].Value = GetScore(scetRecord);
                                        // 平時評量
                                        ws.Cells[5 + studentCount, i + selectSubjectCount * 2 + 1].Value = GetAssignmentScore(scetRecord);
                                    }
                                }


                            }
                        }
                    }

                    // 填總成績 (總成績、排名)
                    if (_awardScoreDict.ContainsKey(sr.ID))
                    {
                        foreach (AwardScore awardScore in _awardScoreDict[sr.ID])
                        {
                            ws.Cells[4, 2 + selectSubjectCount * 4].Value = _examName1 + "總成績";

                            ws.Cells[4, 2 + selectSubjectCount * 4 + 2].Value = _examName2 + "總成績";

                            // 第一次的總成績
                            if (awardScore.ScoreType == "" + ws.Cells[4, 2 + selectSubjectCount * 4].Value)
                            {
                                // 總成績 (除回定期平時的比例100)
                                ws.Cells[5 + studentCount, 2 + selectSubjectCount * 4].Value = awardScore.Score / 100;

                                // 排名
                                ws.Cells[5 + studentCount, 2 + selectSubjectCount * 4 + 1].Value = "QQ";
                            }

                            // 第二次的總成績 (除回定期平時的比例100)
                            if (awardScore.ScoreType == "" + ws.Cells[4, 2 + selectSubjectCount * 4 + 2].Value)
                            {
                                // 總成績
                                ws.Cells[5 + studentCount, 2 + selectSubjectCount * 4 + 2].Value = awardScore.Score / 100;

                                // 排名
                                ws.Cells[5 + studentCount, 2 + selectSubjectCount * 4 + +2 +1].Value = "QQ2";
                            }


                            // 兩次成績差
                            if (awardScore.ScoreType == "進步分數")
                            {
                                // 進步分數 
                                ws.Cells[5 + studentCount, 2 + selectSubjectCount * 4 + 4].Value = awardScore.Score;

                                // 排名
                                ws.Cells[5 + studentCount, 2 + selectSubjectCount * 4 + +4 + 1].Value = "QQ2";
                            }




                        }
                    }




                    studentCount++;

                    _worker.ReportProgress(progress * studentCount, "填寫" + classRecord.Name + "進步獎報表");
                }




                //foreach (Subject subject in term.SubjectList)
                //{
                //    foreach (Assessment assessment in subject.AssessmentList)
                //    {
                //        //分數型成績
                //        if (assessment.Type == "Score")
                //        {


                //            col++;
                //        }
                //    }
                //}

                //// 最後補上 term
                //Cell cell_term = ws.Cells[1, col];
                //cell_term.Copy(wb.Worksheets["樣板一"].Cells["M2"]);

                //cell_term.Value = term.Name;

                //termCol = col;

                //col++;



            }



            ////把多餘的右半邊CELL欄位 砍掉 (總表)             
            //ws_total.Cells.ClearRange(1, progressScoreCol + 1, totalAwardsCount + 2, 50);
            //ws_total.AutoFitColumns();
            //ws_total.FirstVisibleColumn = 0;// 將打開的介面 調到最左， 要不然就會看到 右邊一片空白。


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
            #region 抓評量
            QueryHelper qh = new QueryHelper();

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
            #endregion

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

        // 取得定期分數
        private string GetScore(K12.Data.SCETakeRecord sce)
        {

            XmlElement xmlElement = sce.ToXML();

            XmlElement elem = xmlElement.SelectSingleNode("Extension/Extension/Score") as XmlElement;

            string score = elem == null ? string.Empty : elem.InnerText;

            return score;
        }

        // 取得平時分數
        private string GetAssignmentScore(K12.Data.SCETakeRecord sce)
        {
            XmlElement xmlElement = sce.ToXML();

            XmlElement elem = xmlElement.SelectSingleNode("Extension/Extension/AssignmentScore") as XmlElement;

            string assignmentScore = elem == null ? string.Empty : elem.InnerText;

            return assignmentScore;
        }


        private string GetScoreRatio( K12.Data.AssessmentSetupRecord asr)
        {

            XmlElement xmlElement = asr.Extension;

            XmlElement elem = xmlElement.SelectSingleNode("ScorePercentage") as XmlElement;

            string score = elem == null ? string.Empty : elem.InnerText;

            return score;
        }

    }



}




