using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TalentClassLibrary;
using TalentClassLibrary.Model;
using ShareClassLibrary;
using ServiceStack.Text;
using ServiceStack;
using Newtonsoft;
using Newtonsoft.Json;
using System.IO;

namespace TalentWindowsFormsApp
{
    public partial class ContactStatus : Form
    {
        private bool _CheckChange = false;
        /// <summary>
        /// DB代碼資料
        /// </summary>
        private DataTable CodeDB = new DataTable();
        /// <summary>
        /// UI代碼資料
        /// </summary>
        private List<TextBox> CodeTxt = new List<TextBox>();
        /// <summary>
        /// UI聯繫狀況資料
        /// </summary>
        private DataTable ContactStatusUI = new DataTable();
        /// <summary>
        /// DB聯繫狀況資料
        /// </summary>
        private DataTable ContactStatusDB = new DataTable();
        /// <summary>
        /// DB聯繫基本資料
        /// </summary>
        private DataTable ContactInfoDB = new DataTable();
        /// <summary>
        /// UI聯繫基本資料
        /// </summary>
        private DataTable ContactInfoUI = new DataTable();
        /// <summary>
        /// 聯繫ID
        /// </summary>
        private string Contact_Id = string.Empty;
        /// <summary>
        /// 刪除代碼按鈕
        /// </summary>
        private List<Button> DelCode = new List<Button>();
        /// <summary>
        /// 紀錄目前所在的TabPage
        /// </summary>
        private int Page = 0;
        /// <summary>
        /// 面談資料ID
        /// </summary>
        private List<string> InterviewId = new List<string>();

        /// <summary>
        /// 新增模式
        /// </summary>
        public ContactStatus()
        {
            InitializeComponent();
            CooperationCombo.SelectedItem = "皆可";
            StatusCombo.SelectedItem = "--請選擇--";
            SexCombo.SelectedItem = "--請選擇--";
            ExportCombo.SelectedItem = "聯繫狀況";
            UpdateTimeLbl.Text = "---";

            TextBox code = new TextBox
            {
                Size = new Size(134, 22),
                Location = new Point(98, 42),
                Font = new Font("新細明體", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(136)))
            };
            CodeTxt.Add(code);
            splitContainer2.Panel1.Controls.Add(code);
            ////聯繫狀況欄位
            ContactStatusUI.Columns.Add("Contact_Date");
            ContactStatusUI.Columns.Add("Contact_Status");
            ContactStatusUI.Columns.Add("Remarks");
            PaitDataGridView();
        }

        /// <summary>
        /// 編修模式
        /// </summary>
        /// <param name="contactId">聯繫ID</param>
        public ContactStatus(string contactId)
        {
            InitializeComponent();
            Contact_Id = contactId;
        }

        /// <summary>
        /// 將資料庫的資料顯示出來
        /// </summary>
        private void ShowData()
        {
            DataSet ds = Talent.GetInstance().SelectContactSituationDataById(Contact_Id);
            if (Talent.GetInstance().ErrorMessage != string.Empty)
            {
                MessageBox.Show(Talent.GetInstance().ErrorMessage, "錯誤訊息");
            }

            ContactInfoDB = ds.Tables[0].Copy();
            ContactStatusDB = ds.Tables[1].Copy();
            ContactStatusUI = ds.Tables[1].Copy();
            CodeDB = ds.Tables[2].Copy();
            foreach (DataRow dr in ds.Tables[3].Rows)
            {
                UpdateTimeLbl.Text = dr["UpdateTime"].ToString();
            }
            ////聯繫基本資料
            foreach (DataRow dr in ContactInfoDB.Rows)
            {
                NameTxt.Text = dr["Name"].ToString().Trim();
                SexCombo.SelectedItem = dr["Sex"].ToString().Trim() == string.Empty ? "--請選擇--" : dr["Sex"].ToString().Trim();
                MailTxt.Text = dr["Mail"].ToString().Trim();
                PhoneTxt.Text = dr["CellPhone"].ToString().Trim();
                CooperationCombo.SelectedItem = dr["Cooperation_Mode"].ToString().Trim();
                StatusCombo.SelectedItem = dr["Status"].ToString().Trim() == string.Empty ? "--請選擇--" : dr["Status"].ToString().Trim();
                PlaceTxt.Text = dr["Place"].ToString().Trim();
                SkillTxt.Text = dr["Skill"].ToString().Trim();
                YearTxt.Text = dr["Year"].ToString().Trim();
            }
            ////代碼資料
            if (CodeDB.Rows.Count == 0)
            {
                TextBox code = new TextBox
                {
                    Size = new Size(134, 22),
                    Location = new Point(98, 42),
                    Font = new Font("新細明體", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(136)))
                };
                CodeTxt.Add(code);
                splitContainer2.Panel1.Controls.Add(code);
            }
            else
            {
                for (int i = 0; i < CodeDB.Rows.Count; i++)
                {
                    DataRow dr = CodeDB.Rows[i];
                    TextBox code = new TextBox
                    {
                        Size = new System.Drawing.Size(134, 22),
                        Text = dr["Code_Id"].ToString(),
                        Font = new Font("新細明體", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(136)))
                    };
                    code.Location = i == 0 ? new Point(98, 42) : new Point(98, CodeTxt[CodeTxt.Count - 1].Top + 29);
                    CodeTxt.Add(code);
                    if (i != 0)
                    {
                        Button del = new Button
                        {
                            Size = new Size(49, 19),
                            Image = Image.FromFile("./cancel.png"),
                            Location = new Point(234, code.Top),
                            TabIndex = CodeTxt.Count - 1
                        };
                        del.Click += Del_Click;
                        DelCode.Add(del);
                        splitContainer2.Panel1.Controls.Add(del);
                    }
                    splitContainer2.Panel1.Controls.Add(code);
                }
            }

            CreateInterviewPages(ds.Tables[4]);
        }

        /// <summary>
        /// 根據DB資料創建面談資料TabPages
        /// </summary>
        /// <param name="dt"></param>
        private void CreateInterviewPages(DataTable dt)
        {
            ////在聯繫狀況儲存完，更新頁面時，不要影響面談資料
            if (tabControl1.TabCount > 1)
            {
                return;
            }

            if (dt.Rows.Count == 0)
            {
                //this.tabControl1.TabPages.Add("+");
                return;
            }

            foreach (DataRow dr in dt.Rows)
            {
                InterviewId.Add(dr[0].ToString());
                TabPage tabPage = new TabPage
                {
                    Text = "面談資料" + tabControl1.TabCount,
                    Name = "Interview" + tabControl1.TabCount,
                    AutoScroll = true,
                };
                InterviewDataControl newInterview = new InterviewDataControl(dr[0].ToString(), Contact_Id);
                newInterview.BtnClick += DelInterview_BtnClick;
                newInterview.ConfirmBtnClick += UpdateTime_ConfirmBtnClick;
                tabControl1.TabPages.Add(tabPage);
                tabPage.Controls.Add(newInterview);
            }

            //this.tabControl1.TabPages.Add("+");
        }

        /// <summary>
        /// 新增一個面談資料頁面TabPage
        /// </summary>
        private void NewInterviewPage()
        {
            TabPage tabPage = new TabPage
            {
                Text = "面談資料" + tabControl1.TabCount,
                Name = "Interview" + tabControl1.TabCount,
                AutoScroll = true
            };
            InterviewDataControl newInterview = new InterviewDataControl();
            newInterview.BtnClick += DelInterview_BtnClick;
            newInterview.ConfirmBtnClick += UpdateTime_ConfirmBtnClick;
            newInterview.SetInfo(NameTxt.Text, SexCombo.SelectedItem.ToString(), MailTxt.Text, PhoneTxt.Text, Contact_Id);
            tabControl1.TabPages.Add(tabPage);
            tabPage.Controls.Add(newInterview);
        }

        /// <summary>
        /// 面談資料儲存後更新最後修改時間
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UpdateTime_ConfirmBtnClick(object sender, EventArgs e)
        {
            DataSet ds = Talent.GetInstance().SelectContactSituationDataById(Contact_Id);
            if (!string.IsNullOrEmpty(Talent.GetInstance().ErrorMessage))
            {
                MessageBox.Show(Talent.GetInstance().ErrorMessage, "錯誤訊息");
                return;
            }

            foreach (DataRow dr in ds.Tables[3].Rows)
            {
                UpdateTimeLbl.Text = dr["UpdateTime"].ToString();
            }
        }

        /// <summary>
        /// 刪除面談資料後，重整頁面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DelInterview_BtnClick(object sender, EventArgs e)
        {
            ContactStatus contactStatus = new ContactStatus(Contact_Id)
            {
                MdiParent = this.MdiParent
            };
            contactStatus.Show();
            this.Close();
        }

        /// <summary>
        /// 刪除代碼欄位
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Del_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("是否刪除此代碼?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.No)
            {
                return;
            }

            Button btn = (Button)sender;
            splitContainer2.Panel1.Controls.Remove(CodeTxt[btn.TabIndex]);
            splitContainer2.Panel1.Controls.Remove(DelCode[btn.TabIndex - 1]);
            CodeTxt.RemoveAt(btn.TabIndex);
            DelCode.RemoveAt(btn.TabIndex - 1);
            ////從新定位代碼欄位
            for (int i = 0; i < CodeTxt.Count; i++)
            {
                CodeTxt[i].Location = i == 0 ? new Point(98, 42) : new Point(98, CodeTxt[i - 1].Top + 29);
            }

            for (int i = 0; i < DelCode.Count; i++)
            {
                DelCode[i].Location = new Point(234, CodeTxt[i + 1].Top);
                DelCode[i].TabIndex = CodeTxt.Count - 1;
            }
        }

        /// <summary>
        /// 第一次載入表單時執行
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ContactStatus_Shown(object sender, EventArgs e)
        {
            ImportLink.Visible = false;
            if (!string.IsNullOrEmpty(Contact_Id))
            {
                ShowData();
                PaitDataGridView();
            }
        }

        private void ContactStatus_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "yyyy/MM/dd";
            ExportCombo.SelectedItem = "聯繫狀況";
            this.CreateColumns();
        }

        /// <summary>
        /// 繪製DataGridView
        /// </summary>
        private void PaitDataGridView()
        {
            dataGridView1.Columns.Clear();
            ////在繪製DataGridView時，會觸發CellEnter事件，因此先將其設為唯讀
            dataGridView1.ReadOnly = true;
            dataGridView1.DataSource = ContactStatusUI;
            dataGridView1.Columns[0].HeaderText = "日期";
            dataGridView1.Columns[0].Width = 120;
            dataGridView1.Columns[1].HeaderText = "聯繫狀況";
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].HeaderText = "備註";
            dataGridView1.Columns[2].Width = 1000;
            ////修正案紐-按了之後該列才可編輯
            DataGridViewLinkColumn Update = new DataGridViewLinkColumn
            {
                Text = "修正",
                UseColumnTextForLinkValue = true,
                Width = 60
            };
            dataGridView1.Columns.Add(Update);
            ////刪除案紐-刪除該列
            DataGridViewLinkColumn Del = new DataGridViewLinkColumn
            {
                Text = "刪除",
                UseColumnTextForLinkValue = true,
                Width = 60
            };
            dataGridView1.Columns.Add(Del);
            ////在DataGridView繪製完成後將其設為可編輯模式
            dataGridView1.ReadOnly = false;
            ////將表格設為唯讀
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1.Rows[i].ReadOnly = true;
            }
        }

        /// <summary>
        /// 將控制項定位在DataGridView的Column上
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            this.dateTimePicker1.Visible = false;
            this.ContactStausCombo.Visible = false;
            this.remarkTxt.Visible = false;
            ////只有案修正按鈕才能編輯該列
            if (this.dataGridView1.Rows[e.RowIndex].ReadOnly == true)
            {
                return;
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].HeaderText.Equals("日期"))
            {
                Rectangle r = this.dataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
                dateTimePicker1.Location = new Point(r.X, r.Y);
                this.dateTimePicker1.Size = r.Size;
                this.dateTimePicker1.Height = dataGridView1.Height;
                this._CheckChange = true;
                this.dateTimePicker1.Text = this.dataGridView1.CurrentCell.Value.ToString();
                this._CheckChange = false;
                this.dateTimePicker1.Visible = true;
                this.dateTimePicker1.BringToFront();
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].HeaderText.Equals("聯繫狀況"))
            {
                Rectangle r = this.dataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
                ContactStausCombo.Location = new Point(r.X, r.Y);
                this.ContactStausCombo.Size = r.Size;
                this._CheckChange = true;
                this.ContactStausCombo.Text = this.dataGridView1.CurrentCell.Value.ToString();
                if (string.IsNullOrEmpty(this.ContactStausCombo.Text))
                {
                    this.ContactStausCombo.Text = "(無)";
                }

                this._CheckChange = false;
                this.ContactStausCombo.Visible = true;
                this.ContactStausCombo.BringToFront();
            }

            if (this.dataGridView1.Columns[e.ColumnIndex].HeaderText.Equals("備註"))
            {
                Rectangle r = this.dataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
                remarkTxt.Location = new Point(r.X, r.Y);
                this.remarkTxt.Width = r.Width;
                this._CheckChange = true;
                this.remarkTxt.Text = this.dataGridView1.CurrentCell.Value.ToString();
                this._CheckChange = false;
                this.remarkTxt.Visible = true;
                this.remarkTxt.BringToFront();
            }
        }

        /// <summary>
        /// 聯繫狀況日期
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (_CheckChange) return;
            this.dataGridView1.CurrentCell.Value = this.dateTimePicker1.Text;
        }

        /// <summary>
        /// 聯繫狀況
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ContactStausCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_CheckChange) return;
            this.dataGridView1.CurrentCell.Value = this.ContactStausCombo.SelectedItem;
        }

        /// <summary>
        /// 備註
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void remarkTxt_TextChanged(object sender, EventArgs e)
        {
            if (_CheckChange) return;
            this.dataGridView1.CurrentCell.Value = this.remarkTxt.Text;
        }

        /// <summary>
        /// 新增代碼欄位
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void InsertCodeLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            TextBox code = new TextBox
            {
                Size = new Size(134, 22),
                Location = new Point(98, CodeTxt[CodeTxt.Count - 1].Top + 29),
                Font = new System.Drawing.Font("新細明體", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(136)))
            };
            CodeTxt.Add(code);
            Button del = new Button
            {
                TabIndex = CodeTxt.Count - 1,
                Size = new Size(49, 19),
                Image = Image.FromFile("./cancel.png"),
                Location = new Point(234, code.Top)
            };
            del.Click += Del_Click;
            DelCode.Add(del);
            splitContainer2.Panel1.Controls.Add(code);
            splitContainer2.Panel1.Controls.Add(del);
        }

        /// <summary>
        /// 修正或刪除動作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }

            ////讓該行可以編輯
            if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Equals("修正"))
            {
                dataGridView1.Rows[e.RowIndex].ReadOnly = false;
                dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White;
            }

            ////刪除該列
            if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Equals("刪除"))
            {
                DialogResult result = MessageBox.Show("是否刪除選取的資料", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    ContactStatusUI.Rows.RemoveAt(e.RowIndex);
                }
            }
        }

        /// <summary>
        /// 在上方新增一列
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void InsertContactBtn_Click(object sender, EventArgs e)
        {
            ContactStatusUI.Rows.InsertAt(ContactStatusUI.NewRow(), 0);

            dataGridView1.Rows[0].DefaultCellStyle.BackColor = Color.White;
        }

        /// <summary>
        /// 長紀錄UI資料的DataTable
        /// </summary>
        private void CreateColumns()
        {
            ////聯繫基本資料
            ContactInfoUI.Columns.Add("Name");
            ContactInfoUI.Columns.Add("Sex");
            ContactInfoUI.Columns.Add("Mail");
            ContactInfoUI.Columns.Add("CellPhone");
            ContactInfoUI.Columns.Add("Cooperation_Mode");
            ContactInfoUI.Columns.Add("Status");
            ContactInfoUI.Columns.Add("Place");
            ContactInfoUI.Columns.Add("Skill");
            ContactInfoUI.Columns.Add("Year");
        }

        /// <summary>
        /// 儲存資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ConfirmBtn_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("資料是否正確?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.No)
            {
                return;
            }
            string msg = string.Empty;

            ////UI聯繫基本資料
            ContactInfoUI.Clear();
            string name = NameTxt.Text.Trim();
            string mail = MailTxt.Text.Trim();
            string phone = PhoneTxt.Text.Trim();
            string cooperationMode = CooperationCombo.SelectedItem.ToString();
            string place = PlaceTxt.Text.Trim();
            string skill = SkillTxt.Text.Trim();
            string sex = (SexCombo.SelectedItem.ToString().Equals("--請選擇--")) ? string.Empty : SexCombo.SelectedItem.ToString();
            string status = (StatusCombo.SelectedItem.ToString().Equals("--請選擇--")) ? string.Empty : StatusCombo.SelectedItem.ToString();
            string year = YearTxt.Text;
            ContactInfoUI.Rows.Add(name, sex, mail, phone, cooperationMode, status, place, skill, year);
            ////代碼資料
            DataTable codeList = new DataTable();
            codeList.Columns.Add("Code_Id");
            codeList.Columns.Add("Contact_Id");
            foreach (TextBox code in CodeTxt)
            {
                if (code.Text != string.Empty)
                {
                    codeList.Rows.Add(code.Text, Contact_Id);
                }
            }

            msg = TalentValid.GetInstance().ValidCode(CodeTxt);
            if (msg != string.Empty)
            {
                MessageBox.Show(msg, "錯誤訊息");
                return;
            }

            msg = TalentValid.GetInstance().ValidContactSituationInfoData(name, codeList, sex, mail, phone, place, skill, cooperationMode, status);
            if (msg != string.Empty)
            {
                MessageBox.Show(msg, "錯誤訊息");
                return;
            }

            msg = TalentValid.GetInstance().ValidContactSituationData(ContactStatusUI);
            if (msg != string.Empty)
            {
                MessageBox.Show(msg, "錯誤訊息");
                return;
            }
            ////儲存聯繫基本資料
            if (ContactInfoDB.Rows.Count == 0)
            {
                msg = Talent.GetInstance().InsertContactSituationInfoData(ContactInfoUI);
                if (msg.Equals("新增失敗"))
                {
                    MessageBox.Show(msg, "錯誤訊息");
                    return;
                }
                else
                {
                    ContactInfoDB.Clear();
                    Contact_Id = msg;
                }
            }
            else
            {
                msg = Talent.GetInstance().UpdateContactSituationInfoData(ContactInfoUI, Contact_Id);
                if (msg.Equals("修改失敗"))
                {
                    MessageBox.Show(msg, "錯誤訊息");
                    return;
                }
                else
                {
                    ContactInfoDB.Clear();
                }
            }

            ////儲存代碼資料
            msg = Talent.GetInstance().SaveCode(codeList, Contact_Id);
            if (!msg.Equals("儲存成功"))
            {
                MessageBox.Show(msg, "錯誤訊息");
                return;
            }
            else
            {
                CodeDB.Clear();
                CodeTxt.Clear();
                DelCode.Clear();
            }

            ////儲存聯繫狀況資料
            msg = Talent.GetInstance().SaveContactSituation(ContactStatusUI, Contact_Id);
            if (!msg.Equals("儲存成功"))
            {
                MessageBox.Show(msg, "錯誤訊息");
                return;
            }
            else
            {
                ContactStatusDB.Clear();
                ContactStatusUI.Clear();
            }

            MessageBox.Show("儲存成功", "訊息");
            ShowData();
            PaitDataGridView();
        }

        /// <summary>
        /// 再切換頁面時，如果有改動的資料尚未儲存，則跳出警告
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (this.Page != 0)
            {
                ////面談資料頁面間轉換要檢查是否有存檔
                InterviewDataControl interviewDataControl = new InterviewDataControl();
                TabPage tabPage = tabControl1.TabPages["Interview" + Page];               

                foreach (Control c in tabPage.Controls)
                {
                    if (c is InterviewDataControl)
                    {
                        interviewDataControl = (InterviewDataControl)c;
                        interviewDataControl.InterviewPageChange(e);
                        break;
                    }
                }

                return;
            }

            ////如果聯繫ID等於空則代表沒有此聯繫基本資料
            if (Contact_Id == string.Empty)
            {
                MessageBox.Show("沒有填寫聯繫資料", "警告");
                e.Cancel = true;
                return;
            }

            ////UI聯繫基本資料
            ContactInfo infoUI = new ContactInfo
            {
                Name = NameTxt.Text.Trim(),
                Mail = MailTxt.Text.Trim(),
                CellPhone = PhoneTxt.Text.Trim(),
                Cooperation_Mode = CooperationCombo.SelectedItem.ToString().Trim(),
                Place = PlaceTxt.Text.Trim(),
                Skill = SkillTxt.Text.Trim(),
                Sex = (SexCombo.SelectedItem.ToString().Equals("--請選擇--")) ? string.Empty : SexCombo.SelectedItem.ToString().Trim(),
                Status = (StatusCombo.SelectedItem.ToString().Equals("--請選擇--")) ? string.Empty : StatusCombo.SelectedItem.ToString().Trim(),
                Year = YearTxt.Text
            };
            ////DB聯繫基本資料
            List<ContactInfo> infoDB = ContactInfoDB.DataTableToList<ContactInfo>();
            ////比對是否一樣
            var isdifference = (from DB in infoDB
                                where DB.Name.Trim() == infoUI.Name &&
                                      DB.Mail.Trim() == infoUI.Mail &&
                                      DB.CellPhone.Trim() == infoUI.CellPhone &&
                                      DB.Cooperation_Mode.Trim() == infoUI.Cooperation_Mode &&
                                      DB.Place.Trim() == infoUI.Place &&
                                      DB.Skill.Trim() == infoUI.Skill &&
                                      DB.Sex.Trim() == infoUI.Sex &&
                                      DB.Status.Trim() == infoUI.Status &&
                                      DB.Year.Trim() == infoUI.Year
                                select DB).ToList();
            ////UI代碼資料
            List<string> codeUIList = new List<string>();
            if (CodeTxt.Count == 1 && CodeTxt[0].Text == string.Empty)
            {

            }
            else
            {
                foreach (TextBox code in CodeTxt)
                {
                    codeUIList.Add(code.Text);
                }
            }
            ////DB代碼資料
            List<string> codeDBList = new List<string>();
            foreach (DataRow dr in CodeDB.Rows)
            {
                codeDBList.Add(dr[0].ToString());
            }
            ////UI聯繫狀況資料
            string ContactStatusUIList = JsonConvert.SerializeObject(ContactStatusUI, Formatting.Indented);
            ////UI聯繫狀況資料
            string ContactStatusDBList = JsonConvert.SerializeObject(ContactStatusDB, Formatting.Indented);
            if (isdifference.Count == 0 || !codeDBList.SequenceEqual(codeUIList) || !ContactStatusUIList.Equals(ContactStatusDBList))
            {
                DialogResult result = MessageBox.Show("資料尚未儲存，您確定要離開嗎?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.No)
                {
                    e.Cancel = true;
                }

                return;
            }
        }

        /// <summary>
        /// 清除UI上所有資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DelBtn_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("是否清空此頁面料?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                ////清空聯繫基本資料
                NameTxt.Text = string.Empty;
                MailTxt.Text = string.Empty;
                PhoneTxt.Text = string.Empty;
                CooperationCombo.SelectedItem = "皆可";
                PlaceTxt.Text = string.Empty;
                SkillTxt.Text = string.Empty;
                SexCombo.SelectedItem = "--請選擇--";
                StatusCombo.SelectedItem = "--請選擇--";
                YearTxt.Text = string.Empty;
                ////清空代碼資料
                for (int i = 0; i < CodeTxt.Count; i++)
                {
                    splitContainer2.Panel1.Controls.Remove(CodeTxt[i]);
                }

                for (int i = 0; i < DelCode.Count; i++)
                {
                    splitContainer2.Panel1.Controls.Remove(DelCode[i]);
                }
                CodeTxt.Clear();
                DelCode.Clear();
                TextBox code = new TextBox
                {
                    Size = new System.Drawing.Size(134, 22),
                    Location = new Point(98, 42),
                    Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)))
                };
                CodeTxt.Add(code);
                splitContainer2.Panel1.Controls.Add(code);
                ////清空聯繫狀況資料
                ContactStatusUI.Rows.Clear();
                this.dateTimePicker1.Visible = false;
                this.ContactStausCombo.Visible = false;
                this.remarkTxt.Visible = false;
            }
        }

        /// <summary>
        /// 取得目前的TabPage
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControl1_Deselected(object sender, TabControlEventArgs e)
        {
            Page = e.TabPageIndex;
        }

        /// <summary>
        /// 新增面談資料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NewInterviewBtn_Click(object sender, EventArgs e)
        {
            if (Contact_Id == string.Empty)
            {
                MessageBox.Show("沒有聯繫資料!!!", "警告");
                return;
            }

            this.NewInterviewPage();
        }

        private void ExportLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string path = string.Empty;
            string msg = string.Empty;
            string exportMode = ExportCombo.SelectedItem.ToString();
            if (string.IsNullOrEmpty(Contact_Id))
            {
                MessageBox.Show("沒有聯繫資料，請先儲存資料");
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

            List<string> idList = new List<string>
            {
                Contact_Id
            };

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

        private void ImportLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            if (tabControl1.SelectedIndex == 0)
            {
                return;
            }

            DialogResult result = MessageBox.Show("將會覆蓋此面談資料，是否繼續?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result != DialogResult.Yes)
            {
                return;
            }

            OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "請選擇欲上傳的面談資料Excel",
                Filter = @"Excel Files|*.xlsx",
                Multiselect = false,
                RestoreDirectory = true
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                InterviewDataControl interviewDataControl = new InterviewDataControl();
                TabPage tabPage = tabControl1.TabPages["Interview" + tabControl1.SelectedIndex];
                foreach (Control c in tabPage.Controls)
                {
                    if (c is InterviewDataControl)
                    {
                        interviewDataControl = (InterviewDataControl)c;
                        interviewDataControl.GetdataByExcel(ofd.FileName);
                        break;
                    }
                }
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(tabControl1.SelectedIndex == 0)
            {
                ImportLink.Visible = false;
            }
            else
            {
                ImportLink.Visible = true;
            }
        }
    }
}
