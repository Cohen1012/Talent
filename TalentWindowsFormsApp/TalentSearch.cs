using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TalentClassLibrary;
using TalentClassLibrary.Model;
using ShareClassLibrary;

namespace TalentWindowsFormsApp
{
    public partial class TalentSearch : Form
    {
        public TalentSearch()
        {
            InitializeComponent();
        }

        private void TalentSearch_Load(object sender, EventArgs e)
        {
            CooperationCombo.SelectedItem = "不限";
            ContactStatusCombo.SelectedItem = "不限";
            IsInterviewCombo.SelectedItem = "不限";
            InterviewResultCombo.SelectedItem = "不限";
            ExportCombo.SelectedItem = "聯繫狀況";
            ImportCombo.SelectedItem = "新版";
        }

        private void SearchBtn_Click(object sender, EventArgs e)
        {
            string keyword = KeyWordTxt.Text;
            string skills = SkillsTxt.Text;
            string place = PlaceTxt.Text;
            string cooperationMode = CooperationCombo.SelectedItem.ToString();
            string contactStatus = ContactStatusCombo.SelectedItem.ToString();
            string startEditDate = StartEditDateTxt.Text;
            string endEditDate = EndEditDateTxt.Text;
            string isInterview = IsInterviewCombo.SelectedItem.ToString();
            string interviewResult = InterviewResultCombo.SelectedItem.ToString();
            string startInterviewDate = StartInterviewDateTxt.Text;
            string EndInterviewDate = EndInterviewDateTxt.Text;
            ////從資料庫撈出符合條件的結果
            DataTable dt = TalentClassLibrary.TalentSearch.GetInstance().SelectIdByFilter(keyword, place, skills, cooperationMode, contactStatus, startEditDate, endEditDate, isInterview, interviewResult, startInterviewDate, EndInterviewDate);
            ShowData(ref dt);
        }

        /// <summary>
        /// 將顯示查詢結果
        /// </summary>
        /// <param name="dt"></param>
        private void ShowData(ref DataTable dt)
        {
            if (!string.IsNullOrEmpty(TalentClassLibrary.TalentSearch.GetInstance().ErrorMessage))
            {
                MessageBox.Show(TalentClassLibrary.TalentSearch.GetInstance().ErrorMessage);
                return;
            }

            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("沒有符合的資料", "訊息");
                dataGridView1.Columns.Clear();
                dataGridView1.DataSource = null;
                return;
            }

            dt = TalentClassLibrary.TalentSearch.GetInstance().CombinationGrid(dt);
            if (!string.IsNullOrEmpty(TalentClassLibrary.TalentSearch.GetInstance().ErrorMessage))
            {
                MessageBox.Show(TalentClassLibrary.TalentSearch.GetInstance().ErrorMessage);
                return;
            }

            dataGridView1.Columns.Clear();

            dt.Columns.Add("UpdateTalent");
            dt.Columns.Add("DelTalent");
            foreach (DataRow dr in dt.Rows)
            {
                if (!string.IsNullOrEmpty(dr["Contact_Id"].ToString()))
                {
                    dr["UpdateTalent"] = "修改";
                    dr["DelTalent"] = "刪除";
                }
            }

            ////開始產生Grid
            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.DataSource = dt;

            ////隱藏欄位聯繫ID
            DataGridViewTextBoxColumn contact_Id = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Contact_Id",
                Name = "聯繫ID"
            };
            dataGridView1.Columns.Add(contact_Id);
            ////姓名
            DataGridViewTextBoxColumn name = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Name",
                Name = "姓名",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dataGridView1.Columns.Add(name);
            ////代碼
            DataGridViewTextBoxColumn code = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Code_Id",
                Name = "代碼",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dataGridView1.Columns.Add(code);
            ////聯繫日期
            DataGridViewTextBoxColumn contact_Date = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Contact_Date",
                Name = "聯繫日期",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dataGridView1.Columns.Add(contact_Date);
            ////聯繫狀況
            DataGridViewTextBoxColumn contact_Status = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Contact_Status",
                Name = "聯繫狀況",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dataGridView1.Columns.Add(contact_Status);
            ////備註
            DataGridViewTextBoxColumn remarks = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Remarks",
                Name = "備註",
                Width = 690,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dataGridView1.Columns.Add(remarks);
            ////面試時間
            DataGridViewTextBoxColumn interview_Date = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Interview_Date",
                Name = "面試時間",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dataGridView1.Columns.Add(interview_Date);
            ////最編輯時間
            DataGridViewTextBoxColumn updateTime = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "UpdateTime",
                Name = "最編輯時間",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dataGridView1.Columns.Add(updateTime);
            ////編輯按鈕
            DataGridViewLinkColumn UpdateAccount = new DataGridViewLinkColumn
            {
                DataPropertyName = "UpdateTalent",
                Name = "",
                Width = 50

            };
            dataGridView1.Columns.Add(UpdateAccount);
            ////編輯按鈕
            DataGridViewLinkColumn DelAccount = new DataGridViewLinkColumn
            {
                DataPropertyName = "DelTalent",
                Name = "",
                Width = 50

            };
            dataGridView1.Columns.Add(DelAccount);
            dataGridView1.Columns[0].Visible = false;

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                do
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightSeaGreen;
                    i++;
                    if (i >= dataGridView1.RowCount)
                    {
                        break;
                    }

                } while (string.IsNullOrEmpty(dataGridView1.Rows[i].Cells[0].Value.ToString()));
            }
        }

        private void CalendarBtn_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            CalendarForm calendarForm = new CalendarForm();
            calendarForm.ShowDialog();
            if (calendarForm.DialogResult == DialogResult.OK)
            {
                if (btn.Name.Equals("CalendarBtn"))
                {
                    StartEditDateTxt.Text = calendarForm.CalenderValue;
                }

                if (btn.Name.Equals("CalendarBtn1"))
                {
                    EndEditDateTxt.Text = calendarForm.CalenderValue;
                }

                if (btn.Name.Equals("CalendarBtn2"))
                {
                    StartInterviewDateTxt.Text = calendarForm.CalenderValue;
                }

                if (btn.Name.Equals("CalendarBtn3"))
                {
                    EndInterviewDateTxt.Text = calendarForm.CalenderValue;
                }

                calendarForm.Close();
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }

            ////修改
            if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Equals("修改"))
            {
                ContactStatus contactStatus = new ContactStatus(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString())
                {
                    MdiParent = this.MdiParent
                };
                contactStatus.Show();
                this.Close();
                return;
            }

            ////刪除
            if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Equals("刪除"))
            {
                DialogResult result = MessageBox.Show("是否刪除選取的資料", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    string msg = Talent.GetInstance().DelTalentById(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
                    if (msg.Equals("刪除成功"))
                    {
                        MessageBox.Show(msg, "訊息");
                        dataGridView1.Rows.RemoveAt(e.RowIndex);
                    }
                    else
                    {
                        MessageBox.Show(msg, "錯誤訊息");
                    }
                }
            }
        }

        private void ExportLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string path = string.Empty;
            string msg = string.Empty;
            string exportMode = ExportCombo.SelectedItem.ToString();
            DataTable dt = dataGridView1.DataSource as DataTable;
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("沒有資料", "錯誤訊息");
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog
            {
                Title = "請選擇存檔的路徑",
                FileName = "Excel",
                Filter = @"Excel|*.xlsx"
            };

            if (sfd.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }
            else
            {
                path = Path.GetDirectoryName(sfd.FileName);
            }

            var idList = (from row in dt.AsEnumerable()
                          where !string.IsNullOrEmpty(row.Field<int?>("Contact_Id").ToString())
                          select row.Field<int?>("Contact_Id").ToString()).ToList();

            switch (exportMode)
            {
                case "聯繫狀況":
                    msg = TalentCommon.GetInstance().ExportContactSituationDataByContactId(idList, path);
                    break;
                case "面談資料":
                    msg = TalentCommon.GetInstance().ExportInterviewDataByContactId(idList, path);
                    break;
                case "所有資料":
                    msg = TalentCommon.GetInstance().ExportAllDataByContactId(idList, path);
                    break;
                default:
                    msg = "沒有匯出任何資料";
                    break;
            }

            MessageBox.Show(msg, "訊息");

        }

        private void InsertContactBtn_Click(object sender, EventArgs e)
        {
            ContactStatus contactStatus = new ContactStatus
            {
                MdiParent = this.MdiParent
            };
            contactStatus.Show();
            this.Close();
        }

        private void ImportLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string importMode = ImportCombo.SelectedItem.ToString();
            List<ContactSituation> contactSituationList = new List<ContactSituation>();
            OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "請選擇上傳的聯繫狀況Excel",
                Filter = @"Excel Files|*.xlsx",
                Multiselect = false,
                RestoreDirectory = true
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                switch (importMode)
                {
                    case "新版":
                        contactSituationList = ExcelHelper.GetInstance().ImportNewTalent(ofd.FileName);
                        break;
                    case "舊版":
                        contactSituationList = ExcelHelper.GetInstance().ImportOldTalent(ofd.FileName);
                        break;
                    default:
                        MessageBox.Show("沒有匯入任何資料","錯誤訊息");
                        break;
                }

                if (!string.IsNullOrEmpty(ExcelHelper.GetInstance().ErrorMessage))
                {
                    MessageBox.Show(ExcelHelper.GetInstance().ErrorMessage);
                    return;
                }

                DataTable dt = TalentClassLibrary.TalentSearch.GetInstance().SelectTop15();
                ShowData(ref dt);
                MessageBox.Show(TalentClassLibrary.TalentSearch.GetInstance().InsertContactSituationInfoData(contactSituationList));
            }
        }

        /// <summary>
        /// 第一次載入表單時執行
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TalentSearch_Shown(object sender, EventArgs e)
        {
            DataTable dt = TalentClassLibrary.TalentSearch.GetInstance().SelectTop15();
            ShowData(ref dt);
        }
    }
}
