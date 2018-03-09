using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TalentClassLibrary.Model;
using Newtonsoft.Json;
using TalentClassLibrary;
using ShareClassLibrary;
using ServiceStack.Text;
using ServiceStack;

namespace TalentWindowsFormsApp
{
    public partial class InterviewDataControl : UserControl
    {
        /// <summary>
        /// 面談資料ID
        /// </summary>
        public string Interview_Id = string.Empty;

        /// <summary>
        /// 觸發事件
        /// </summary>
        public event EventHandler BtnClick;

        /// <summary>
        /// 觸發儲存事件
        /// </summary>
        public event EventHandler ConfirmBtnClick;

        /// <summary>
        /// 觸發更新最後修改時間事件
        /// </summary>
        /// <param name="e"></param>
        protected void UpdateTimeClick(EventArgs e)
        {
            ConfirmBtnClick?.Invoke(this, e);
        }

        /// <summary>
        /// 觸發刪除面談資料事件
        /// </summary>
        /// <param name="e"></param>
        protected void DelInterviewDataClick(EventArgs e)
        {
            BtnClick?.Invoke(this, e);
        }

        /// <summary>
        /// 聯繫ID
        /// </summary>
        private string Contact_Id = string.Empty;

        /// <summary>
        /// DB面談資料
        /// </summary>
        private DataTable InterviewInfoDB = new DataTable();

        /// <summary>
        /// UI學歷資料
        /// </summary>
        private DataTable EducationUI = new DataTable();

        /// <summary>
        /// UI工作經驗
        /// </summary>
        private DataTable WorkExperienceUI = new DataTable();

        /// <summary>
        /// UI語言能力
        /// </summary>
        private DataTable LanguageUI = new DataTable();

        /// <summary>
        /// UI面談評語
        /// </summary>
        private DataTable InterviewCommentsUI = new DataTable();

        /// <summary>
        /// DB面談評語
        /// </summary>
        private DataTable InterviewCommentsDB = new DataTable();

        /// <summary>
        /// DB面談結果
        /// </summary>
        private DataTable InterviewResultDB = new DataTable();

        /// <summary>
        /// DB專案經驗
        /// </summary>
        private DataTable ProjectExperienceDB = new DataTable();

        /// <summary>
        /// UI專案經驗
        /// </summary>
        private List<ProjectExperienceControl> ProjectExperienceUI = new List<ProjectExperienceControl>();

        /// <summary>
        /// 紀錄目前所在的TabPage
        /// </summary>
        private int Page = 0;

        private bool EducationCheckChange = false;
        private bool WorkExperienceCheckChange = false;
        private bool LanguageCheckChange = false;
        private bool CommentCheckChange = false;

        /// <summary>
        /// 將基本資訊傳入面談資料
        /// </summary>
        /// <param name="name">姓名</param>
        /// <param name="sex">性別</param>
        /// <param name="mail">e-mail</param>
        /// <param name="phone">手機</param>
        public void SetInfo(string name, string sex, string mail, string phone, string contact_Id)
        {
            NameTxt.Text = name;
            SexCombo.SelectedItem = sex;
            MailTxt.Text = mail;
            PhoneTxt.Text = phone;
            Contact_Id = contact_Id;
        }

        /// <summary>
        /// 從DB讀取資料
        /// </summary>
        public void GetdataByDB()
        {
            DataSet ds = Talent.GetInstance().SelectInterviewDataById(Interview_Id);
            if (!string.IsNullOrEmpty(Talent.GetInstance().ErrorMessage))
            {
                MessageBox.Show(Talent.GetInstance().ErrorMessage, "錯誤訊息");
                return;
            }

            this.UpdateInterviewPage(ds);
            this.UpdateInterviewResult(ds);
            this.UpdateProjectExperience(ds);
        }

        /// <summary>
        /// 從excel讀取資料
        /// </summary>
        /// <param name="path">檔案路徑</param>
        public void GetdataByExcel(string path)
        {
            DataSet ds = ExcelHelper.GetInstance().ImportInterviewData(path);
            if (!string.IsNullOrEmpty(ExcelHelper.GetInstance().ErrorMessage))
            {
                MessageBox.Show(ExcelHelper.GetInstance().ErrorMessage, "錯誤訊息");
                return;
            }

            ////先將DB資料備份
            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(InterviewInfoDB.Copy());
            dataSet.Tables.Add(InterviewCommentsDB.Copy());
            dataSet.Tables.Add(InterviewResultDB.Copy());
            dataSet.Tables.Add(ProjectExperienceDB.Copy());
            ////清空畫面
            InterviewInfoDB.Clear();
            ProjectExperienceDB.Clear();
            ProjectExperienceUI.Clear();
            panel2.Controls.Clear();         

            UpdateInterviewPage(ds);
            UpdateProjectExperience(ds);
            ////要保留原本DB的資料
            InterviewInfoDB = dataSet.Tables[0].Copy();           
            ProjectExperienceDB = dataSet.Tables[3].Copy();

            MessageBox.Show("匯入成功");
        }

        /// <summary>
        /// 面談資料間轉換，要卡是否有儲存資料
        /// </summary>
        /// <param name="e"></param>
        public void InterviewPageChange(TabControlCancelEventArgs e)
        {
            ////面談基本資料頁面
            if (this.Page == 0)
            {
                this.CompareInterviewInfo(e);
            }

            ////面談結果頁面
            if (this.Page == 2)
            {
                this.CompareInterviewResult(e);
            }

            ////專案經驗頁面
            if (this.Page == 1)
            {
                this.CompareProjectExperience(e);
            }
        }

        public InterviewDataControl()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 編修模式
        /// </summary>
        /// <param name="interviewId"></param>
        /// <param name="contactId"></param>
        public InterviewDataControl(string interviewId, string contactId)
        {
            InitializeComponent();
            Interview_Id = interviewId;
            Contact_Id = contactId;
            this.GetdataByDB();
        }

        /// <summary>
        /// 選取時間
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CalendarBtn_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            CalendarForm calendarForm = new CalendarForm();
            calendarForm.ShowDialog();
            if (calendarForm.DialogResult == DialogResult.OK)
            {
                BirthdayTxt.Text = calendarForm.CalenderValue;
                calendarForm.Close();
            }
        }

        private void InterviewDataControl_Load(object sender, EventArgs e)
        {
            if (Interview_Id == string.Empty)
            {
                NewInterviewPage();
            }
        }

        /// <summary>
        /// 面談資料修改模式
        /// </summary>
        private void UpdateInterviewPage(DataSet ds)
        {
            InterviewInfoDB = ds.Tables[0].Copy();
            InterviewInfoDB = this.DBNullToEmpty(InterviewInfoDB);
            this.ShowInterviewInfo(InterviewInfoDB);
        }

        /// <summary>
        /// 專案經驗
        /// </summary>
        private void UpdateProjectExperience(DataSet ds)
        {           
            ProjectExperienceDB = ds.Tables[3].Copy();
            ProjectExperienceDB = this.DBNullToEmpty(ProjectExperienceDB);
            this.ShowProjectExperience();
        }

        /// <summary>
        /// 面談結果
        /// </summary>
        private void UpdateInterviewResult(DataSet ds)
        {           
            InterviewCommentsDB = ds.Tables[1].Copy();
            InterviewCommentsDB = this.DBNullToEmpty(InterviewCommentsDB);
            InterviewCommentsUI = ds.Tables[1].Copy();
            InterviewResultDB = ds.Tables[2].Copy();
            InterviewResultDB = this.DBNullToEmpty(InterviewResultDB);
            this.ShowInterviewResult();
        }

        /// <summary>
        /// 將DBNull轉成String.Empty
        /// </summary>
        /// <param name="dataTable"></param>
        /// <returns></returns>
        private DataTable DBNullToEmpty(DataTable dataTable)
        {
            foreach (DataRow dr in dataTable.Rows)
            {
                foreach (DataColumn dc in dataTable.Columns)
                {
                    if (DBNull.Value.Equals(dr[dc]))
                    {
                        dr[dc] = string.Empty;
                    }
                }
            }

            return dataTable;
        }

        /// <summary>
        /// 定位專案經驗
        /// </summary>
        private void RepositionProjectExperience()
        {
            for (int i = ProjectExperienceUI.Count - 1; i >= 0; i--)
            {
                ProjectExperienceUI[i].Location = new Point(38, 20 + (((ProjectExperienceUI.Count - 1 - i) * 65) + (ProjectExperienceUI[i].Height * ((ProjectExperienceUI.Count - 1) - i))));
            }
        }

        /// <summary>
        /// 顯示專案資料
        /// </summary>
        private void ShowProjectExperience()
        {
            List<ProjectExperience> projectExperienceList = ProjectExperienceDB.DataTableToList<ProjectExperience>();
            for (int i = 0; i < projectExperienceList.Count; i++)
            {
                ProjectExperienceControl projectExperienceControl = new ProjectExperienceControl();
                projectExperienceControl.BtnClick += DelProjectExperienceControl_BtnClick;
                ProjectExperienceUI.Add(projectExperienceControl);
                panel2.Controls.Add(projectExperienceControl);
                projectExperienceControl.SetProjectExperience(projectExperienceList[i]);
            }

            this.RepositionProjectExperience();
        }

        /// <summary>
        /// 移除指定專案經驗
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DelProjectExperienceControl_BtnClick(object sender, EventArgs e)
        {
            ProjectExperienceControl projectExperienceControl = (ProjectExperienceControl)sender;
            ProjectExperienceUI.Remove(projectExperienceControl);
            panel2.Controls.Remove(projectExperienceControl);
            this.RepositionProjectExperience();
        }

        /// <summary>
        /// 顯示面談結果資料
        /// </summary>
        private void ShowInterviewResult()
        {
            foreach (DataRow dr in InterviewResultDB.Rows)
            {
                ////任用評定
                HireRadio.Checked = false;
                NoHireRadio.Checked = false;
                ReserveRadio.Checked = false;
                if (dr["Appointment"].ToString().Trim().Equals("錄用"))
                {
                    HireRadio.Checked = true;
                }

                if (dr["Appointment"].ToString().Trim().Equals("不錄用"))
                {
                    NoHireRadio.Checked = true;
                }

                if (dr["Appointment"].ToString().Trim().Equals("暫保留"))
                {
                    ReserveRadio.Checked = true;
                }
                ////備註
                CommentRemarkTxt.Text = dr["Results_Remark"].ToString().Trim();
            }

            PaitInterviewResultGrid();
        }

        /// <summary>
        /// 繪製面評語Grid
        /// </summary>
        private void PaitInterviewResultGrid()
        {
            CommentsGrid.DataSource = null;
            CommentsGrid.Columns.Clear();
            CommentsGrid.DataSource = InterviewCommentsUI;
            CommentsGrid.Columns["Interviewer"].HeaderText = "面談者";
            CommentsGrid.Columns["Interviewer"].Width = 100;
            CommentsGrid.Columns["Result"].HeaderText = "面談結果";
            CommentsGrid.Columns["Result"].Width = 1000;
            CommentsGrid.Columns.Add(DelColumn());
        }

        /// <summary>
        /// 顯示面談基本資料
        /// </summary>
        /// <param name="interviewInfoDB"></param>
        private void ShowInterviewInfo(DataTable interviewInfoDB)
        {
            foreach (DataRow dr in interviewInfoDB.Rows)
            {
                VacanciesTxt.Text = dr["Vacancies"].ToString().Trim();
                InterviewDateTxt.Text = dr["Interview_Date"].ToString().Trim();
                NameTxt.Text = dr["Name"].ToString().Trim();
                SexCombo.SelectedItem = dr["Sex"].ToString().Trim() == string.Empty ? "--請選擇--" : dr["Sex"].ToString().Trim();
                BirthdayTxt.Text = dr["Birthday"].ToString().Trim();
                MarriedCombo.SelectedItem = dr["Married"].ToString().Trim() == string.Empty ? "--請選擇--" : dr["Married"].ToString().Trim();
                MailTxt.Text = dr["Mail"].ToString().Trim();
                PhoneTxt.Text = dr["CellPhone"].ToString().Trim();
                AdressTxt.Text = dr["Adress"].ToString().Trim();
                pictureBox1.ImageLocation = dr["Image"].ToString().Trim();
                UrgentContactPersonTxt.Text = dr["Urgent_Contact_Person"].ToString().Trim();
                UrgentRelationshipTxt.Text = dr["Urgent_Relationship"].ToString().Trim();
                UrgentCellPhoneTxt.Text = dr["Urgent_CellPhone"].ToString().Trim();
                DuringServiceTxt.Text = dr["During_Service"].ToString().Trim();
                ExemptionReasonTxt.Text = dr["Exemption_Reason"].ToString().Trim();
                IsStudyTxt.Text = dr["IsStudy"].ToString().Trim();
                IsServiceTxt.Text = dr["IsService"].ToString().Trim();
                RelativesRelationshipTxt.Text = dr["Relatives_Relationship"].ToString().Trim();
                RelativesNameTxt.Text = dr["Relatives_Name"].ToString().Trim();
                CareWorkTxt.Text = dr["Care_Work"].ToString().Trim();
                HopeSalaryTxt.Text = dr["Hope_Salary"].ToString().Trim();
                WhenReportTxt.Text = dr["When_Report"].ToString().Trim();
                AdvantageTxt.Text = dr["Advantage"].ToString().Trim();
                DisAdvantageTxt.Text = dr["Disadvantages"].ToString().Trim();
                HobbyTxt.Text = dr["Hobby"].ToString().Trim();
                AttractReasonTxt.Text = dr["Attract_Reason"].ToString().Trim();
                FutureGoalTxt.Text = dr["Future_Goal"].ToString().Trim();
                HopeSupervisorTxt.Text = dr["Hope_Supervisor"].ToString().Trim();
                HopePromiseTxt.Text = dr["Hope_Promise"].ToString().Trim();
                IntroductionTxt.Text = dr["Introduction"].ToString().Trim();
                ////學歷
                if (dr["Education"].ToString().Trim() == string.Empty)
                {
                    this.CreateEducationColumns();
                }
                else
                {
                    EducationUI = JsonConvert.DeserializeObject<DataTable>(dr["Education"].ToString().Trim());
                }

                ////工作經驗
                if (dr["Work_Experience"].ToString().Trim() == string.Empty)
                {
                    this.CreateWorkExperienceColumns();
                }
                else
                {
                    WorkExperienceUI = JsonConvert.DeserializeObject<DataTable>(dr["Work_Experience"].ToString().Trim());
                }

                ////語言能力
                if (dr["Language"].ToString().Trim() == string.Empty)
                {
                    this.CreateLanguageColumns();
                }
                else
                {
                    LanguageUI = JsonConvert.DeserializeObject<DataTable>(dr["Language"].ToString().Trim());
                }

                ////程式語言
                this.ChcekedExpertise(dr["Expertise_Language"].ToString().Trim(), languageCheckedListBox, languageTxt);
                ////開發工具
                this.ChcekedExpertise(dr["Expertise_Tools"].ToString().Trim(), ToolsCheckedListBox, ToolsTxt);
                FramworkTxt.Text = dr["Expertise_Tools_Framwork"].ToString().Trim();
                ////Devops
                this.ChcekedExpertise(dr["Expertise_Devops"].ToString().Trim(), DevopsCheckedListBox, DevopsTxt);
                ////作業系統
                this.ChcekedExpertise(dr["Expertise_OS"].ToString().Trim(), OSCheckedListBox, OSTxt);
                ////大數據
                this.ChcekedExpertise(dr["Expertise_BigData"].ToString().Trim(), BigDataCheckedListBox, BigDataTxt);
                ////資料庫
                this.ChcekedExpertise(dr["Expertise_DataBase"].ToString().Trim(), DataBaseCheckedListBox, DataBaseTxt);
                ////專業認證
                this.ChcekedExpertise(dr["Expertise_Certification"].ToString().Trim(), CertificationCheckedListBox, CertificationTxt);
                PaitDataGridView();
            }
        }

        /// <summary>
        /// 勾選專長
        /// </summary>
        /// <param name="expertiseList">專長字串以","隔開</param>
        /// <param name="expertiseCheckedListBox">專長的CheckedListBox物件</param>
        /// <param name="expertiseText">專長的其他欄位</param>
        private void ChcekedExpertise(string expertiseList, CheckedListBox expertiseCheckedListBox, TextBox expertiseText)
        {
            if (!string.IsNullOrEmpty(expertiseList))
            {
                expertiseText.Text = string.Empty;
                string[] expertise = expertiseList.Split(',');
                for (int i = 0; i < expertise.Length; i++)
                {
                    for (int j = 0; j < expertiseCheckedListBox.Items.Count; j++)
                    {
                        if (expertiseCheckedListBox.Items[j].ToString() == expertise[i])
                        {
                            expertiseCheckedListBox.SetItemChecked(j, true);
                            break;
                        }

                        if (j == expertiseCheckedListBox.Items.Count - 1)
                        {
                            expertiseText.Text += expertise[i] + ",";
                        }
                    }
                }

                expertiseText.Text = expertiseText.Text.RemoveEndWithDelimiter(",");
            }
        }

        /// <summary>
        /// 學歷欄位
        /// </summary>
        private void CreateEducationColumns()
        {
            if (EducationUI.Columns.Count > 0)
            {
                return;
            }

            EducationUI.Columns.Add("School");
            EducationUI.Columns.Add("Department");
            EducationUI.Columns.Add("Start_End_Date");
            EducationUI.Columns.Add("Is_Graduation");
            EducationUI.Columns.Add("Remark");
        }

        /// <summary>
        /// 工作經驗欄位
        /// </summary>
        private void CreateWorkExperienceColumns()
        {
            if (WorkExperienceUI.Columns.Count > 0)
            {
                return;
            }

            WorkExperienceUI.Columns.Add("Institution_name");
            WorkExperienceUI.Columns.Add("Position");
            WorkExperienceUI.Columns.Add("Start_End_Date");
            WorkExperienceUI.Columns.Add("Start_Salary");
            WorkExperienceUI.Columns.Add("End_Salary");
            WorkExperienceUI.Columns.Add("Leaving_Reason");
        }

        /// <summary>
        /// 語言能力欄位
        /// </summary>
        private void CreateLanguageColumns()
        {
            if (LanguageUI.Columns.Count > 0)
            {
                return;
            }

            LanguageUI.Columns.Add("Language_Name");
            LanguageUI.Columns.Add("Listen");
            LanguageUI.Columns.Add("Speak");
            LanguageUI.Columns.Add("Read");
            LanguageUI.Columns.Add("Write");
        }

        /// <summary>
        /// 面談資料新增模式
        /// </summary>
        private void NewInterviewPage()
        {
            MarriedCombo.SelectedItem = "--請選擇--";
            HireRadio.Checked = false;
            NoHireRadio.Checked = false;
            ReserveRadio.Checked = false;
            this.CreateEducationColumns();
            this.CreateWorkExperienceColumns();
            this.CreateLanguageColumns();
            PaitDataGridView();
        }

        /// <summary>
        /// 繪製DataGridView
        /// </summary>
        private void PaitDataGridView()
        {
            ////學歷
            EducationGrid.DataSource = null;
            EducationGrid.Columns.Clear();
            EducationGrid.DataSource = EducationUI;
            EducationGrid.Columns["School"].HeaderText = "學校名稱";
            EducationGrid.Columns["School"].Width = 200;
            EducationGrid.Columns["Department"].HeaderText = "科系";
            EducationGrid.Columns["Department"].Width = 200;
            EducationGrid.Columns["Start_End_Date"].HeaderText = "起迄年月";
            EducationGrid.Columns["Start_End_Date"].Width = 200;
            EducationGrid.Columns["Is_Graduation"].HeaderText = "畢/肄業";
            EducationGrid.Columns["Is_Graduation"].Width = 100;
            EducationGrid.Columns["Remark"].HeaderText = "備註";
            EducationGrid.Columns["Remark"].Width = 500;
            EducationGrid.Columns.Add(DelColumn());
            ////工作經驗
            WorkExperienceGrid.DataSource = null;
            WorkExperienceGrid.Columns.Clear();
            WorkExperienceGrid.DataSource = WorkExperienceUI;
            WorkExperienceGrid.Columns["Institution_name"].HeaderText = "機構名稱";
            WorkExperienceGrid.Columns["Institution_name"].Width = 150;
            WorkExperienceGrid.Columns["Position"].HeaderText = "職稱";
            WorkExperienceGrid.Columns["Position"].Width = 150;
            WorkExperienceGrid.Columns["Start_End_Date"].HeaderText = "起迄年月";
            WorkExperienceGrid.Columns["Start_End_Date"].Width = 150;
            WorkExperienceGrid.Columns["Start_Salary"].HeaderText = "到職薪資";
            WorkExperienceGrid.Columns["Start_Salary"].Width = 150;
            WorkExperienceGrid.Columns["End_Salary"].HeaderText = "離職薪資";
            WorkExperienceGrid.Columns["End_Salary"].Width = 150;
            WorkExperienceGrid.Columns["Leaving_Reason"].HeaderText = "離職原因";
            WorkExperienceGrid.Columns["Leaving_Reason"].Width = 500;
            WorkExperienceGrid.Columns.Add(DelColumn());
            ////語言能力
            LanguageGrid.DataSource = null;
            LanguageGrid.Columns.Clear();
            LanguageGrid.DataSource = LanguageUI;
            LanguageGrid.Columns["Language_Name"].HeaderText = "語言";
            LanguageGrid.Columns["Language_Name"].Width = 100;
            LanguageGrid.Columns["Listen"].HeaderText = "聽";
            LanguageGrid.Columns["Listen"].Width = 100;
            LanguageGrid.Columns["Speak"].HeaderText = "說";
            LanguageGrid.Columns["Speak"].Width = 100;
            LanguageGrid.Columns["Read"].HeaderText = "讀";
            LanguageGrid.Columns["Read"].Width = 100;
            LanguageGrid.Columns["Write"].HeaderText = "寫";
            LanguageGrid.Columns["Write"].Width = 100;
            LanguageGrid.Columns.Add(DelColumn());
        }

        /// <summary>
        /// 加入刪除欄位
        /// </summary>
        /// <returns></returns>
        private DataGridViewLinkColumn DelColumn()
        {
            ////刪除案紐-刪除該列
            DataGridViewLinkColumn del = new DataGridViewLinkColumn
            {
                Text = "刪除",
                UseColumnTextForLinkValue = true
            };
            return del;
        }

        /// <summary>
        /// 將控制項定位在學歷表格的Column上
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void EducationGrid_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            this.GraduationCombo.Visible = false;
            this.EducationRemarkTxt.Visible = false;
            if (this.EducationGrid.Columns[e.ColumnIndex].HeaderText.Equals( "備註"))
            {
                Rectangle r = this.EducationGrid.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
                EducationRemarkTxt.Location = new Point(r.X, r.Y);
                this.EducationRemarkTxt.Size = r.Size;
                this.EducationRemarkTxt.Height = EducationGrid.Height;
                this.EducationCheckChange = true;
                this.EducationRemarkTxt.Text = this.EducationGrid.CurrentCell.Value.ToString();
                this.EducationCheckChange = false;
                this.EducationRemarkTxt.Visible = true;
                this.EducationRemarkTxt.BringToFront();
            }

            if (this.EducationGrid.Columns[e.ColumnIndex].HeaderText.Equals("畢/肄業"))
            {
                Rectangle r = this.EducationGrid.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
                GraduationCombo.Location = new Point(r.X, r.Y);
                this.GraduationCombo.Size = r.Size;
                this.GraduationCombo.Height = EducationGrid.Height;
                this.EducationCheckChange = true;
                this.GraduationCombo.Text = string.IsNullOrEmpty(this.EducationGrid.CurrentCell.Value.ToString()) ? "畢業" : this.EducationGrid.CurrentCell.Value.ToString();
                this.EducationCheckChange = false;
                this.GraduationCombo.Visible = true;
                this.GraduationCombo.BringToFront();
            }
        }

        /// <summary>
        /// 將控制項定位在工作經驗表格的Column上
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void WorkExperienceGrid_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            this.LeavingReasonTxt.Visible = false;
            if (this.WorkExperienceGrid.Columns[e.ColumnIndex].HeaderText.Equals("離職原因"))
            {
                Rectangle r = this.WorkExperienceGrid.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
                LeavingReasonTxt.Location = new Point(r.X, r.Y);
                this.LeavingReasonTxt.Size = r.Size;
                this.LeavingReasonTxt.Height = WorkExperienceGrid.Height;
                this.WorkExperienceCheckChange = true;
                this.LeavingReasonTxt.Text = this.WorkExperienceGrid.CurrentCell.Value.ToString();
                this.WorkExperienceCheckChange = false;
                this.LeavingReasonTxt.Visible = true;
                this.LeavingReasonTxt.BringToFront();
            }
        }

        /// <summary>
        /// 將控制項定位在語言能力表格的Column上
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LanguageGrid_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            this.AbilityCombo.Visible = false;
            if (LanguageGrid.Columns[e.ColumnIndex].HeaderText.Equals("聽") || LanguageGrid.Columns[e.ColumnIndex].HeaderText.Equals("說") || LanguageGrid.Columns[e.ColumnIndex].HeaderText.Equals("讀") || LanguageGrid.Columns[e.ColumnIndex].HeaderText.Equals("寫"))
            {
                Rectangle r = LanguageGrid.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
                AbilityCombo.Location = new Point(r.X, r.Y);
                this.AbilityCombo.Size = r.Size;
                this.AbilityCombo.Height = LanguageGrid.Height;
                this.LanguageCheckChange = true;
                this.AbilityCombo.Text = string.IsNullOrEmpty(LanguageGrid.CurrentCell.Value.ToString()) ? "略" : LanguageGrid.CurrentCell.Value.ToString();

                this.LanguageCheckChange = false;
                this.AbilityCombo.Visible = true;
                this.AbilityCombo.BringToFront();
            }
        }

        /// <summary>
        /// 將控制項定位在面談評語表格的Column上
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CommentsGrid_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            this.CommentTxt.Visible = false;
            if (CommentsGrid.Columns[e.ColumnIndex].HeaderText.Equals("面談結果"))
            {
                Rectangle r = CommentsGrid.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
                CommentTxt.Location = new Point(r.X, r.Y);
                this.CommentTxt.Size = r.Size;
                this.CommentTxt.Height = LanguageGrid.Height;
                this.CommentCheckChange = true;
                this.CommentTxt.Text = CommentsGrid.CurrentCell.Value.ToString();
                this.CommentCheckChange = false;
                this.CommentTxt.Visible = true;
                this.CommentTxt.BringToFront();
            }
        }

        /// <summary>
        /// 學歷-備註
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void EducationRemarkTxt_TextChanged(object sender, EventArgs e)
        {
            if (EducationCheckChange) return;
            this.EducationGrid.CurrentCell.Value = this.EducationRemarkTxt.Text;
        }

        /// <summary>
        /// 學歷-畢/肄業
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GraduationCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (EducationCheckChange) return;
            this.EducationGrid.CurrentCell.Value = this.GraduationCombo.SelectedItem.ToString();
        }

        /// <summary>
        /// 面談評語-面談結果
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CommentTxt_TextChanged(object sender, EventArgs e)
        {
            if (CommentCheckChange) return;
            this.CommentsGrid.CurrentCell.Value = this.CommentTxt.Text;
        }

        /// <summary>
        /// 工作經驗-離職原因
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LeavingReasonTxt_TextChanged(object sender, EventArgs e)
        {
            if (WorkExperienceCheckChange) return;
            this.WorkExperienceGrid.CurrentCell.Value = this.LeavingReasonTxt.Text;
        }

        /// <summary>
        /// 語言能力
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AbilityCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (LanguageCheckChange) return;
            this.LanguageGrid.CurrentCell.Value = this.AbilityCombo.SelectedItem.ToString();
        }

        /// <summary>
        /// 刪除該列學歷
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void EducationGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }

            if (EducationGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Equals("刪除"))
            {
                DialogResult result = MessageBox.Show("是否刪除選取的資料", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    EducationUI.Rows.RemoveAt(e.RowIndex);
                }
            }
        }

        /// <summary>
        /// 刪除該列工作經驗
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void WorkExperienceGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }

            if (WorkExperienceGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Equals("刪除"))
            {
                DialogResult result = MessageBox.Show("是否刪除選取的資料", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    WorkExperienceUI.Rows.RemoveAt(e.RowIndex);
                }
            }
        }

        /// <summary>
        /// 刪除該列語言能力
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LanguageGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }

            if (LanguageGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Equals("刪除"))
            {
                DialogResult result = MessageBox.Show("是否刪除選取的資料", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    LanguageUI.Rows.RemoveAt(e.RowIndex);
                }
            }
        }

        /// <summary>
        /// 刪除該列面談評語
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CommentsGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }

            if (CommentsGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Equals("刪除"))
            {
                DialogResult result = MessageBox.Show("是否刪除選取的資料", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    InterviewCommentsUI.Rows.RemoveAt(e.RowIndex);
                }
            }
        }

        /// <summary>
        /// 學歷新增一列
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void InEducationBtn_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            EducationUI.Rows.Add(EducationUI.NewRow());
        }

        /// <summary>
        /// 工作經驗新增一列
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void InWorkExperienceBtn_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            WorkExperienceUI.Rows.Add(WorkExperienceUI.NewRow());
        }

        /// <summary>
        /// 語言能力新增一列
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void InLanguageBtn_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LanguageUI.Rows.Add(LanguageUI.NewRow());
        }

        /// <summary>
        /// 面談評語新增一列
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void InsertCommentsBtn_Click(object sender, EventArgs e)
        {
            InterviewCommentsUI.Rows.Add(InterviewCommentsUI.NewRow());
        }

        /// <summary>
        /// 新增一個專案經驗
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NewProjectExperienceBtn_Click(object sender, EventArgs e)
        {
            ProjectExperienceControl projectExperienceControl = new ProjectExperienceControl();
            projectExperienceControl.BtnClick += DelProjectExperienceControl_BtnClick;
            ProjectExperienceUI.Add(projectExperienceControl);
            panel2.Controls.Add(projectExperienceControl);

            this.RepositionProjectExperience();

        }

        /// <summary>
        /// 將專長CheckBox串成字串
        /// </summary>
        /// <param name="checkedListBox">checkedListBox物件</param>
        /// <param name="expertiseTxt">專長的其他欄位</param>
        /// <returns></returns>
        private string GetExpertise(CheckedListBox checkedListBox, string expertiseTxt)
        {
            string expertise = string.Empty;
            for (int i = 0; i < checkedListBox.Items.Count; i++)
            {
                if (checkedListBox.GetItemChecked(i))
                {
                    expertise += checkedListBox.Items[i].ToString() + ",";
                }
            }

            expertise += expertiseTxt;

            expertise = expertise.RemoveEndWithDelimiter(",");

            ////去除重複的專長
            expertise = string.Join(",", expertise.Split(',').Distinct().ToArray());

            return expertise;
        }

        /// <summary>
        /// 將面談基本資料頁面包成類別
        /// </summary>
        /// <returns></returns>
        private InterviewInfo CreateUIInterviewInfo()
        {
            InterviewInfo interviewInfoUI = new InterviewInfo
            {
                Vacancies = VacanciesTxt.Text,
                Interview_Date = InterviewDateTxt.Value.ToString("yyyy/MM/dd"),
                Name = NameTxt.Text,
                Sex = (SexCombo.SelectedItem.ToString().Equals("--請選擇--")) ? string.Empty : SexCombo.SelectedItem.ToString(),
                Birthday = BirthdayTxt.Text,
                Married = (MarriedCombo.SelectedItem.ToString().Equals("--請選擇--")) ? string.Empty : MarriedCombo.SelectedItem.ToString(),
                Mail = MailTxt.Text,
                CellPhone = PhoneTxt.Text,
                Adress = AdressTxt.Text,
                //Image = string.IsNullOrEmpty(pictureBox1.ImageLocation) ? "" : pictureBox1.ImageLocation,
                Urgent_Contact_Person = UrgentContactPersonTxt.Text,
                Urgent_Relationship = UrgentRelationshipTxt.Text,
                Urgent_CellPhone = UrgentCellPhoneTxt.Text,
                Education = JsonConvert.SerializeObject(EducationUI, Formatting.Indented),
                Work_Experience = JsonConvert.SerializeObject(WorkExperienceUI, Formatting.Indented),
                Language = JsonConvert.SerializeObject(LanguageUI, Formatting.Indented),
                During_Service = DuringServiceTxt.Text,
                Exemption_Reason = ExemptionReasonTxt.Text,
                IsStudy = IsStudyTxt.Text,
                IsService = IsServiceTxt.Text,
                Relatives_Relationship = RelativesRelationshipTxt.Text,
                Relatives_Name = RelativesNameTxt.Text,
                Care_Work = CareWorkTxt.Text,
                Hope_Salary = HopeSalaryTxt.Text,
                When_Report = WhenReportTxt.Text,
                Advantage = AdvantageTxt.Text,
                Disadvantages = DisAdvantageTxt.Text,
                Hobby = HobbyTxt.Text,
                Attract_Reason = AttractReasonTxt.Text,
                Future_Goal = FutureGoalTxt.Text,
                Hope_Supervisor = HopeSupervisorTxt.Text,
                Hope_Promise = HopePromiseTxt.Text,
                Introduction = IntroductionTxt.Text,
                ////程式語言
                Expertise_Language = this.GetExpertise(languageCheckedListBox, languageTxt.Text),
                ////開發工具
                Expertise_Tools = this.GetExpertise(ToolsCheckedListBox, ToolsTxt.Text),
                ////開發工具-Framwork
                Expertise_Tools_Framwork = FramworkTxt.Text,
                ////Devops
                Expertise_Devops = this.GetExpertise(DevopsCheckedListBox, DevopsTxt.Text),
                ////作業系統
                Expertise_OS = this.GetExpertise(OSCheckedListBox, OSTxt.Text),
                ////大數據
                Expertise_BigData = this.GetExpertise(BigDataCheckedListBox, BigDataTxt.Text),
                ////資料庫
                Expertise_DataBase = this.GetExpertise(DataBaseCheckedListBox, DataBaseTxt.Text),
                ////專業認證
                Expertise_Certification = this.GetExpertise(CertificationCheckedListBox, CertificationTxt.Text)
            };
            ////教育程度
            interviewInfoUI.Education = interviewInfoUI.Education.Equals("[]") ? string.Empty : interviewInfoUI.Education;
            ////工作經歷
            interviewInfoUI.Work_Experience = interviewInfoUI.Work_Experience.Equals("[]" )? string.Empty : interviewInfoUI.Work_Experience;
            ////語言能力
            interviewInfoUI.Language = interviewInfoUI.Language.Equals("[]") ? string.Empty : interviewInfoUI.Language;

            return interviewInfoUI;
        }

        /// <summary>
        /// 儲存面談資料
        /// </summary>
        private void SaveInterviewData(EventArgs e)
        {
            string msg = string.Empty;
            string validMsg = string.Empty;
            DialogResult result = MessageBox.Show("資料是否正確?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.No)
            {
                return;
            }

            ////驗證專案經驗資料
            List<ProjectExperience> projectExperienceUIList = new List<ProjectExperience>();
            foreach (ProjectExperienceControl projectExperienceUI in ProjectExperienceUI)
            {
                ProjectExperience project = new ProjectExperience();
                project = projectExperienceUI.GetProjectExperience();
                projectExperienceUIList.Add(project);
            }

            if (TalentValid.GetInstance().ValidProjectExperienceData(projectExperienceUIList) != string.Empty)
            {
                DialogResult result1 = MessageBox.Show("有空的專案經驗是否儲存?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result1 == DialogResult.No)
                {
                    return;
                }
            }

            validMsg = TalentValid.GetInstance().ValidProjectExperienceIsRepeat(projectExperienceUIList.ListToDataTable());
            if (validMsg != string.Empty)
            {
                msg += validMsg + "\n";
            }

            ////驗證面談基本資料
            InterviewInfo interviewInfoUI = this.CreateUIInterviewInfo();
            interviewInfoUI.Image = string.IsNullOrEmpty(pictureBox1.ImageLocation) ? string.Empty : pictureBox1.ImageLocation;
            validMsg = TalentValid.GetInstance().ValidInterviewInfoData(interviewInfoUI.Vacancies,
                interviewInfoUI.Name, interviewInfoUI.Married, interviewInfoUI.Interview_Date,
                interviewInfoUI.Sex, interviewInfoUI.Mail, interviewInfoUI.Birthday,
                interviewInfoUI.CellPhone, interviewInfoUI.Image);
            List<InterviewInfo> interviewInfoUIList = new List<InterviewInfo>
            {
                   interviewInfoUI
            };

            DataTable dataTable = interviewInfoUIList.ListToDataTable();
            if (validMsg != string.Empty)
            {
                msg += validMsg;
            }

            ////驗證面談結果資料
            InterviewResult interviewResultUI = CreateUIInterviewResult();
            List<InterviewResult> interviewResultUIList = new List<InterviewResult>
            {
                interviewResultUI
            };

            if (string.IsNullOrEmpty(interviewResultUI.Appointment) && InterviewCommentsUI.Rows.Count > 0)
            {
                msg += "請選擇任用評定\n";
            }

            if (msg != string.Empty)
            {
                MessageBox.Show(msg, "錯誤訊息");
                return;
            }

            DataSet ds = new DataSet();
            ds.Tables.Add(InterviewCommentsUI.Copy());
            ds.Tables.Add(interviewInfoUIList.ListToDataTable().Copy());
            ds.Tables.Add(projectExperienceUIList.ListToDataTable().Copy());
            ds.Tables.Add(interviewResultUIList.ListToDataTable().Copy());

            ////儲存面談資料
            msg = TalentClassLibrary.TalentSearch.GetInstance().SaveInterviewData(ds, Interview_Id, InterviewInfoDB.DataTableToList<InterviewInfo>()[0].Image, interviewInfoUI.Image);

            if (!msg.Equals("修改成功"))
            {
                MessageBox.Show(msg, "錯誤訊息");
                return;
            }
            else
            {
                MessageBox.Show("儲存成功", "訊息");
                UpdateTimeClick(e);

                InterviewInfoDB.Clear();
                ProjectExperienceDB.Clear();
                ProjectExperienceUI.Clear();
                panel2.Controls.Clear();
                InterviewResultDB.Clear();
                InterviewCommentsDB.Clear();
                this.GetdataByDB();
            }
        }

        /// <summary>
        /// 儲存面談基本資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void InfoConfirmBtn_Click(object sender, EventArgs e)
        {
            ////新增面談基本資料
            if (string.IsNullOrEmpty(Interview_Id))
            {
                DialogResult result = MessageBox.Show("資料是否正確?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.No)
                {
                    return;
                }

                string msg = string.Empty;
                InterviewInfo interviewInfoUI = this.CreateUIInterviewInfo();
                interviewInfoUI.Image = string.IsNullOrEmpty(pictureBox1.ImageLocation) ? string.Empty : pictureBox1.ImageLocation;

                msg = TalentValid.GetInstance().ValidInterviewInfoData(interviewInfoUI.Vacancies,
                    interviewInfoUI.Name, interviewInfoUI.Married, interviewInfoUI.Interview_Date,
                    interviewInfoUI.Sex, interviewInfoUI.Mail, interviewInfoUI.Birthday,
                    interviewInfoUI.CellPhone, interviewInfoUI.Image);
                if (msg != string.Empty)
                {
                    MessageBox.Show(msg, "錯誤訊息");
                    return;
                }

                List<InterviewInfo> interviewInfoUIList = new List<InterviewInfo>
                {
                   interviewInfoUI
                };
                DataTable dataTable = interviewInfoUIList.ListToDataTable();

                msg = Talent.GetInstance().InsertInterviewInfoData(interviewInfoUIList.ListToDataTable(), Contact_Id);
                if (msg.Equals("新增失敗"))
                {
                    MessageBox.Show(msg, "錯誤訊息");
                    return;
                }
                else
                {
                    InterviewInfoDB.Clear();
                    Interview_Id = msg;
                }

                MessageBox.Show("儲存成功", "訊息");

                UpdateTimeClick(e);

                DataSet dataSet = Talent.GetInstance().SelectInterviewDataById(Interview_Id);
                if (!string.IsNullOrEmpty(Talent.GetInstance().ErrorMessage))
                {
                    MessageBox.Show(Talent.GetInstance().ErrorMessage, "錯誤訊息");
                    return;
                }
                this.UpdateInterviewPage(dataSet);
                // this.UpdateInterviewResult(dataSet);
            }
            else
            {
                this.SaveInterviewData(e);
            }
        }

        /// <summary>
        /// 儲存專案經驗
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ProjectConfirmBtn_Click(object sender, EventArgs e)
        {
            this.SaveInterviewData(e);           
        }

        /// <summary>
        /// 儲存面談結果
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ResultConfirmBtn_Click(object sender, EventArgs e)
        {
            this.SaveInterviewData(e);          
        }

        /// <summary>
        /// 將面談結果頁面包成類別
        /// </summary>
        /// <returns></returns>
        private InterviewResult CreateUIInterviewResult()
        {
            InterviewResult interviewResultUI = new InterviewResult();
            interviewResultUI.Appointment = string.Empty;
            if (HireRadio.Checked)
            {
                interviewResultUI.Appointment = HireRadio.Text;
            }

            if (NoHireRadio.Checked)
            {
                interviewResultUI.Appointment = NoHireRadio.Text;
            }

            if (ReserveRadio.Checked)
            {
                interviewResultUI.Appointment = ReserveRadio.Text;
            }

            interviewResultUI.Results_Remark = CommentRemarkTxt.Text;
            return interviewResultUI;
        }

        /// <summary>
        /// 上傳圖片
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UpLoadBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "請選擇上傳的圖片",

                Filter = @"All Image Files|*.bmp;*.ico;*.gif;*.jpeg;*.jpg;*.png;*.tif;*.tiff",

                Multiselect = false,
                RestoreDirectory = true
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                pictureBox1.ImageLocation = ofd.FileName;
            }
        }

        /// <summary>
        /// 移除照片
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RemoveBtn_Click(object sender, EventArgs e)
        {
            pictureBox1.ImageLocation = null;
        }

        /// <summary>
        /// 清空專案經驗畫面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DelProjectBtn_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("是否清空此頁面料?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                ProjectExperienceUI.Clear();
                panel2.Controls.Clear();
            }
        }

        /// <summary>
        /// 清空面談結果頁面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DelResultBtn_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("是否清空此頁面料?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                HireRadio.Checked = true;
                CommentRemarkTxt.Text = string.Empty;
                InterviewCommentsUI.Rows.Clear();
                CommentTxt.Visible = false;
            }
        }

        /// <summary>
        /// 清空面談基本資料頁面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DelBtn_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("是否清空此頁面料?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                VacanciesTxt.Text = string.Empty;
                InterviewDateTxt.Text = string.Empty;
                NameTxt.Text = string.Empty;
                SexCombo.SelectedItem = "--請選擇--";
                BirthdayTxt.Text = string.Empty;
                MarriedCombo.SelectedItem = "--請選擇--";
                MailTxt.Text = string.Empty;
                PhoneTxt.Text = string.Empty;
                AdressTxt.Text = string.Empty;
                pictureBox1.ImageLocation = string.Empty;
                UrgentContactPersonTxt.Text = string.Empty;
                UrgentRelationshipTxt.Text = string.Empty;
                UrgentCellPhoneTxt.Text = string.Empty;
                DuringServiceTxt.Text = string.Empty;
                ExemptionReasonTxt.Text = string.Empty;
                IsStudyTxt.Text = string.Empty;
                IsServiceTxt.Text = string.Empty;
                RelativesRelationshipTxt.Text = string.Empty;
                RelativesNameTxt.Text = string.Empty;
                CareWorkTxt.Text = string.Empty;
                HopeSalaryTxt.Text = string.Empty;
                WhenReportTxt.Text = string.Empty;
                AdvantageTxt.Text = string.Empty;
                DisAdvantageTxt.Text = string.Empty;
                HobbyTxt.Text = string.Empty;
                AttractReasonTxt.Text = string.Empty;
                FutureGoalTxt.Text = string.Empty;
                HopeSupervisorTxt.Text = string.Empty;
                HopePromiseTxt.Text = string.Empty;
                IntroductionTxt.Text = string.Empty;
                pictureBox1.ImageLocation = null;
                ////學歷
                EducationUI.Rows.Clear();
                GraduationCombo.Visible = false;
                EducationRemarkTxt.Visible = false;
                ////工作經驗
                WorkExperienceUI.Rows.Clear();
                LeavingReasonTxt.Visible = false;
                ////語言能力
                LanguageUI.Rows.Clear();
                AbilityCombo.Visible = false;
                ////程式語言
                for (int i = 0; i < languageCheckedListBox.Items.Count; i++)
                {
                    languageCheckedListBox.SetItemChecked(i, false);
                }
                languageTxt.Text = string.Empty;
                ////開發工具
                for (int i = 0; i < ToolsCheckedListBox.Items.Count; i++)
                {
                    ToolsCheckedListBox.SetItemChecked(i, false);
                }
                FramworkTxt.Text = string.Empty;
                ToolsTxt.Text = string.Empty;
                ////Devops
                for (int i = 0; i < DevopsCheckedListBox.Items.Count; i++)
                {
                    DevopsCheckedListBox.SetItemChecked(i, false);
                }
                DevopsTxt.Text = string.Empty;
                ////作業系統
                for (int i = 0; i < OSCheckedListBox.Items.Count; i++)
                {
                    OSCheckedListBox.SetItemChecked(i, false);
                }
                OSTxt.Text = string.Empty;
                ////大數據
                for (int i = 0; i < BigDataCheckedListBox.Items.Count; i++)
                {
                    BigDataCheckedListBox.SetItemChecked(i, false);
                }
                BigDataTxt.Text = string.Empty;
                ////資料庫
                for (int i = 0; i < DataBaseCheckedListBox.Items.Count; i++)
                {
                    DataBaseCheckedListBox.SetItemChecked(i, false);
                }
                DataBaseTxt.Text = string.Empty;
                ////專業認證
                for (int i = 0; i < CertificationCheckedListBox.Items.Count; i++)
                {
                    CertificationCheckedListBox.SetItemChecked(i, false);
                }
                CertificationTxt.Text = string.Empty;
            }
        }

        /// <summary>
        /// 紀錄目前的TabPage
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControl1_Deselected(object sender, TabControlEventArgs e)
        {
            this.RepositionProjectExperience();
            Page = e.TabPageIndex;
        }

        /// <summary>
        /// 再切換頁面時，如果有改動的資料尚未儲存，則跳出警告
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            ////面談基本資料頁面
            if (this.Page == 0)
            {
                ////如果面談ID等於空則代表沒有此面談基本資料
                if (Interview_Id == string.Empty)
                {
                    MessageBox.Show("沒有填寫面談基本資料", "警告");
                    e.Cancel = true;
                    return;
                }

                this.CompareInterviewInfo(e);
            }

            ////面談結果頁面
            if (this.Page == 2)
            {
                this.CompareInterviewResult(e);
            }

            ////專案經驗頁面
            if (this.Page == 1)
            {
                this.CompareProjectExperience(e);
            }

            if (!e.Cancel)
            {
                this.Page = tabControl1.SelectedIndex;
            }
        }

        /// <summary>
        /// 比對專案經驗頁面是否有變動
        /// </summary>
        private void CompareProjectExperience(TabControlCancelEventArgs e)
        {
            List<ProjectExperience> projectExperienceUIList = new List<ProjectExperience>();
            foreach (ProjectExperienceControl projectExperienceUI in ProjectExperienceUI)
            {
                ProjectExperience project = new ProjectExperience();
                project = projectExperienceUI.GetProjectExperience();
                projectExperienceUIList.Add(project);
            }

            string ProjectExperienceUIJson = JsonConvert.SerializeObject(projectExperienceUIList, Formatting.Indented);
            string ProjectExperienceDBJson = JsonConvert.SerializeObject(ProjectExperienceDB, Formatting.Indented);
            if (!ProjectExperienceDBJson.Equals(ProjectExperienceUIJson))
            {
                DialogResult result = MessageBox.Show("專案經驗資料尚未儲存，您確定要離開嗎?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.No)
                {
                    e.Cancel = true;
                }

                return;
            }
        }

        /// <summary>
        /// 比對面談結果頁面是否有變動
        /// </summary>
        private void CompareInterviewResult(TabControlCancelEventArgs e)
        {
            InterviewResult interviewResultUI = CreateUIInterviewResult();
            List<InterviewResult> interviewResultUIList = new List<InterviewResult>
            {
                interviewResultUI
            };
            string interviewResultUIJson = JsonConvert.SerializeObject(interviewResultUIList, Formatting.Indented);
            string interviewResultDBJson = JsonConvert.SerializeObject(InterviewResultDB, Formatting.Indented);

            string interviewCommentUIJson = JsonConvert.SerializeObject(InterviewCommentsUI, Formatting.Indented);
            string interviewCommentDBJson = JsonConvert.SerializeObject(InterviewCommentsDB, Formatting.Indented);
            if (!interviewResultUIJson.Equals(interviewResultDBJson) || !interviewCommentUIJson.Equals(interviewCommentDBJson))
            {
                DialogResult result = MessageBox.Show("面談結果資料尚未儲存，您確定要離開嗎?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.No)
                {
                    e.Cancel = true;
                }

                return;
            }
        }

        /// <summary>
        /// 比對面談基本資料是否有變動
        /// </summary>
        private void CompareInterviewInfo(TabControlCancelEventArgs e)
        {
            InterviewInfo interviewInfoUI = this.CreateUIInterviewInfo();
            interviewInfoUI.Image = string.IsNullOrEmpty(pictureBox1.ImageLocation) ? "" : pictureBox1.ImageLocation;
            List<InterviewInfo> interviewInfoUIList = new List<InterviewInfo>
            {
                interviewInfoUI
            };
            string interviewInfoUIJson = JsonConvert.SerializeObject(interviewInfoUIList.ListToDataTable(), Formatting.Indented);
            string interviewInfoDBJson = JsonConvert.SerializeObject(InterviewInfoDB, Formatting.Indented);

            if (!interviewInfoUIJson.Equals(interviewInfoDBJson))
            {
                DialogResult result = MessageBox.Show("面談基本資料資料尚未儲存，您確定要離開嗎?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.No)
                {
                    e.Cancel = true;
                }

                return;
            }
        }

        /// <summary>
        /// 刪除此份面談資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DelInterviewDataBtn_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("是否刪除此份面談資料", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.No)
            {
                return;
            }

            if (Interview_Id == string.Empty)
            {
                DelInterviewDataClick(e);
            }

            if (Interview_Id != string.Empty)
            {
                string msg = Talent.GetInstance().DelInterviewDataByInterviewId(Interview_Id);
                if (!msg.Equals("刪除成功"))
                {
                    MessageBox.Show(msg, "錯誤訊息");
                    return;
                }

                MessageBox.Show("刪除成功", "訊息");
                DelInterviewDataClick(e);
            }
        }

        private void AttachedFiles_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LinkLabel btn = sender as LinkLabel;
            if (string.IsNullOrEmpty(Interview_Id))
            {
                MessageBox.Show("請先建立面談基本資料", "警告");
                return;
            }

            string attachedFilesMode = string.Empty;

            switch (btn.Name)
            {
                case "ProjectFilesLink":
                    {
                        attachedFilesMode = "ProjectExperience";
                        break;
                    }
                case "InterViewInfoFilesLink":
                    {
                        attachedFilesMode = "InterviewInfo";
                        break;
                    }
            }

            Files files = new Files(Interview_Id, attachedFilesMode);
            if (files.ShowDialog() == DialogResult.OK)
            {
                this.UpdateTimeClick(e);
                files.Close();
            }
        }
    }
}
