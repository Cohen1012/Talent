using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TalentClassLibrary;
using Microsoft.VisualBasic;
using ShareClassLibrary;

namespace TalentWindowsFormsApp
{
    public partial class AuthorityManagement : Form
    {
        /// <summary>
        /// 帳號
        /// </summary>
        public string Account { get; set; }

        public AuthorityManagement(string account)
        {
            InitializeComponent();
            Account = account;
        }

        /// <summary>
        /// 帳號管理
        /// </summary>
        /// <param name="account"></param>
        private void Manage_Accounts(string account)
        {
            if (account.Equals("hr@is-land.com.tw"))
            {
                InsertAccountBtn.Visible = true;
                dataGridView2.Visible = true;

                dataGridView2.AutoGenerateColumns = false;
                DataTable dt = Talent.GetInstance().SelectMemberInfo();
                if (!string.IsNullOrEmpty(Talent.GetInstance().ErrorMessage))
                {
                    MessageBox.Show(Talent.GetInstance().ErrorMessage);
                    return;
                }

                dataGridView2.DataSource = dt;
                dt.Columns.Add("Reset");
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["States"].ToString().Equals("啟用"))
                    {
                        dr["Reset"] = "重設密碼";
                    }
                }

                dataGridView2.Columns.Clear();
                ////帳號欄位               
                DataGridViewTextBoxColumn accountColumn = new DataGridViewTextBoxColumn
                {
                    DataPropertyName = "Account",
                    Name = "帳號",
                    ReadOnly = true
                };
                dataGridView2.Columns.Add(accountColumn);
                ////狀態欄位
                dataGridView2.Columns.Add(CreateComboBox());
                ////重設密碼欄位
                DataGridViewLinkColumn resetPassword = new DataGridViewLinkColumn
                {
                    DataPropertyName = "Reset",
                    Name = "重設密碼"
                };
                dataGridView2.Columns.Add(resetPassword);
                ////刪除帳號欄位
                DataGridViewColumn delAccount = new DataGridViewButtonColumn
                {
                    Text = "刪除",
                    Name = "刪除",
                    UseColumnTextForButtonValue = true
                };
                dataGridView2.Columns.Add(delAccount);
            }
            else
            {
                InsertAccountBtn.Visible = false;
                dataGridView2.Visible = false;
            }
        }

        /// <summary>
        /// 連結狀態欄位
        /// </summary>
        /// <returns></returns>
        private DataGridViewComboBoxColumn CreateComboBox()
        {
            DataGridViewComboBoxColumn states = new DataGridViewComboBoxColumn
            {
                DataPropertyName = "States",
                Name = "狀態",
                DataSource = new string[] { "啟用", "停用" }
            };
            return states;
        }

        private void AuthorityManagement_Load(object sender, EventArgs e)
        {
            DataTable dt = Talent.GetInstance().SelectMemberInfoByAccount(Account);
            if (!string.IsNullOrEmpty(Talent.GetInstance().ErrorMessage))
            {
                MessageBox.Show(Talent.GetInstance().ErrorMessage);
                return;
            }

            dt.Columns[0].ColumnName = "帳號";
            dt.Columns[1].ColumnName = "狀態";
            dataGridView1.DataSource = dt;
            DataGridViewColumn updatePassword = new DataGridViewLinkColumn
            {
                HeaderText = "修改密碼",
                Text = "修改密碼",
                UseColumnTextForLinkValue = true
            };
            dataGridView1.Columns.Insert(dataGridView1.Columns.Count, updatePassword);
            this.Manage_Accounts(Account);
        }

        /// <summary>
        /// 修改密碼
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            UpdatePassword updatePassword = new UpdatePassword
            {
                Account = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString()
            };
            this.DialogResult = updatePassword.ShowDialog();
            if (this.DialogResult == DialogResult.OK)
            {
                MessageBox.Show("密碼修改成功，請重新登入");
                Application.Restart();
            }
            else if (this.DialogResult == DialogResult.Cancel)
            {
                updatePassword.Close();
            }
        }
        /// <summary>
        /// 新增帳號
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void InsertAccountBtn_Click(object sender, EventArgs e)
        {
            string newAccount = Interaction.InputBox("請輸入新帳號", "新增帳號", null);
            if (newAccount != null)
            {
                string msg = Talent.GetInstance().InsertMember(newAccount);
                if (!msg.Equals("新增成功"))
                {
                    MessageBox.Show(msg, "錯誤訊息");
                }
                else
                {
                    MessageBox.Show(msg + "\n密碼預設為帳號", "訊息");
                    this.Manage_Accounts(Account);
                }
            }
        }

        /// <summary>
        /// 根據動作執行"更改狀態"or"重設密碼"or"刪除帳號"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }

            ////重設密碼
            if (dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Equals("重設密碼"))
            {
                for (int i = 0; i < dataGridView2.ColumnCount; i++)
                {
                    if (dataGridView2.Rows[e.RowIndex].Cells[i].Value.ToString().Contains("@"))
                    {
                        string msg = TalentCommon.GetInstance().AlertUpdatePassword(dataGridView2.Rows[e.RowIndex].Cells[i].Value.ToString());
                        if (!msg.Equals("寄送成功"))
                        {
                            MessageBox.Show(msg, "錯誤訊息");
                        }
                        else
                        {
                            MessageBox.Show("密碼重設成功，請去信箱收信", "訊息");
                        }
                        return;
                    }
                }
            }
            ////刪除帳號
            if (dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Equals("刪除"))
            {
                DialogResult result = MessageBox.Show("即將刪除該帳號是否繼續?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    for (int i = 0; i < dataGridView2.ColumnCount; i++)
                    {
                        if (dataGridView2.Rows[e.RowIndex].Cells[i].Value.ToString().Contains("@"))
                        {
                            string msg = Talent.GetInstance().DelMemberByAccount(dataGridView2.Rows[e.RowIndex].Cells[i].Value.ToString());
                            if (!msg.Equals("刪除成功"))
                            {
                                MessageBox.Show(msg, "錯誤訊息");
                            }
                            else
                            {
                                MessageBox.Show(msg, "訊息");
                                this.Manage_Accounts(Account);
                            }
                        }
                    }
                }
            }
        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }

            ////更改狀態
            if (dataGridView2.Columns[e.ColumnIndex].Name.ToString().Equals("狀態"))
            {
                string account = dataGridView2.Rows[e.RowIndex].Cells["帳號"].Value.ToString();
                string states = dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                string msg = Talent.GetInstance().UpdateMemberStatesByAccount(account, states);
                MessageBox.Show(msg, "訊息");
                this.Manage_Accounts(Account);
            }
        }


        private void dataGridView2_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //ComboBox combo = e.Control as ComboBox;
            //if (combo != null)
            //{
            //    ////Remove an existing event-handler, if present, to avoid 
            //    ////adding multiple handlers when the editing control is reused.
            //    combo.SelectedIndexChanged -= Combo_SelectedIndexChanged;
            //    combo.SelectedIndexChanged += Combo_SelectedIndexChanged;
            //}
        }

        private void Combo_SelectedIndexChanged(object sender, EventArgs e)
        {
            //ComboBox combo = sender as ComboBox;
            //MessageBox.Show(combo.SelectedItem.ToString());
        }

        private void dataGridView2_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView2.IsCurrentCellDirty)
            {
                dataGridView2.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }
    }
}