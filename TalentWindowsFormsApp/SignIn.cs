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

namespace TalentWindowsFormsApp
{
    public partial class SignIn : Form
    {
        public string Msg { get; set; }

        public SignIn()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 登入動作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SignInBtn_Click(object sender, EventArgs e)
        {
            string account = AccountTxt.Text;
            string password = PasswordTxt.Text;
            Msg = Talent.GetInstance().SignIn(account, password);
            if (!string.IsNullOrEmpty(Talent.GetInstance().ErrorMessage))
            {
                MessageBox.Show(Talent.GetInstance().ErrorMessage);
                return;
            }

            if (!Msg.Equals("登入失敗") && !Msg.Equals("該帳號停用中"))
            {
                this.DialogResult = DialogResult.OK;
                this.Hide();
            }
            else
            {
                MessageBox.Show(Msg, "錯誤訊息");
            }
        }

        /// <summary>
        /// 帳號欄位按下Enter，焦點換到密碼去
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AccountTxt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                PasswordTxt.Focus();
            }
        }

        /// <summary>
        /// 密碼欄位按下Enter後，執行登入動作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PasswordTxt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                this.SignInBtn_Click(sender, e);
            }
        }
    }
}
