using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TalentWindowsFormsApp
{
    public partial class UpdatePassword : Form
    {
        public string Account { get; set; }
        public UpdatePassword()
        {
            InitializeComponent();
        }

        private void UpdatePassword_Load(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 確認按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OKBtn_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("即將修改密碼是否繼續?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                string oldPassword = OldPasswordTxt.Text;
                string newPassword = NewPasswordTxt.Text;
                string checkNewPassword = CheckNewPasswordTxt.Text;
                string msg = TalentClassLibrary.Talent.GetInstance().UpdatePasswordByaccount(Account, oldPassword, newPassword, checkNewPassword);
                if (!msg.Equals("修改成功"))
                {
                    MessageBox.Show(msg, "錯誤訊息");
                }
                else
                {
                    this.DialogResult = DialogResult.OK;
                    this.Hide();
                }
            }
        }

        /// <summary>
        /// 取消按鈕
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CancelBtn_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("是否取消修改密碼?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                this.DialogResult = DialogResult.Cancel;
                this.Hide();
            }
        }

        /// <summary>
        /// 按Enter將焦點移至新密碼欄位
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OldPasswordTxt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(e.KeyChar == 13)
            {
                NewPasswordTxt.Focus();
            }
        }

        /// <summary>
        /// 按Enter將焦點移至確認新密碼欄位
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NewPasswordTxt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                CheckNewPasswordTxt.Focus();
            }
        }

        /// <summary>
        /// 按Enter觸發修改密碼動作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckNewPasswordTxt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                this.OKBtn_Click(sender, e);
            }
        }
    }
}
