using ShareClassLibrary;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TalentClassLibrary.Model;
using Spire.Xls;

namespace TalentClassLibrary
{
    /// <summary>
    /// 驗證資料類別
    /// </summary>
    public partial class TalentValid
    {
        private static TalentValid talentValid = new TalentValid();

        public static TalentValid GetInstance() => talentValid;

        /// <summary>
        /// 驗證聯繫狀況資料的正確性
        /// </summary>
        /// <param name="validData">聯繫狀況資料</param>
        /// <returns>空值代表沒有錯誤</returns>
        public string ValidContactSituationData(DataTable validData)
        {
            string msg = string.Empty;
            foreach (DataRow dr in validData.Rows)
            {
                msg = Valid.GetInstance().ValidDateFormat(dr["Contact_Date"].ToString().Trim());
                if (!msg.Equals(string.Empty))
                {
                    return msg;
                }

                msg = this.ValidContactStatus(dr["Contact_Status"].ToString().Trim());
                if (!msg.Equals(string.Empty))
                {
                    return msg;
                }
            }

            return msg;
        }

        /// <summary>
        /// 驗證聯繫狀況基本資料是否正確
        /// </summary>
        /// <param name="name">姓名不可為空</param>
        /// <param name="code">可以有多個代碼，代碼都是唯一值</param>
        /// <param name="sex">值為"男"or"女"</param>
        /// <param name="mail">e-mail格式</param>
        /// <param name="cellPhone">手機格式</param>
        /// <param name="place"><地點</param>
        /// <param name="skill">技能</param>
        /// <param name="cooperationMode">合作模式值為"全職"or"合約"or"皆可"</param>
        /// <param name="states">狀態值為"追蹤"or"保留"</param>
        /// <returns>空值代表沒有錯誤</returns>
        public string ValidContactSituationInfoData(string name, DataTable code, string sex, string mail, string cellPhone, string place, string skill, string cooperationMode, string states)
        {
            string msg = string.Empty;
            string validMsg = string.Empty;
            if (string.IsNullOrEmpty(name) && code.Rows.Count == 0)
            {
                msg += "姓名或代碼請至少填一個\n";
            }

            validMsg = Valid.GetInstance().ValidSex(sex);
            if (!validMsg.Equals(string.Empty))
            {
                msg += validMsg + "\n";
            }

            validMsg = this.ValidCooperationMode(cooperationMode);
            if (!validMsg.Equals(string.Empty))
            {
                msg += validMsg + "\n";
            }

            validMsg = this.ValidStates(states);
            if (!validMsg.Equals(string.Empty))
            {
                msg += validMsg + "\n";
            }

            validMsg = Talent.GetInstance().ValidCodeIsRepeat(code);
            if (!validMsg.Equals(string.Empty))
            {
                msg += validMsg;
            }

            return msg;
        }

        /// <summary>
        /// 驗證合作模式的值，值為"全職"or"合約"or"皆可"
        /// </summary>
        /// <param name="cooperationMode">驗證模式的值</param>
        /// <returns>空值代表沒有錯誤</returns>
        public string ValidCooperationMode(string cooperationMode)
        {
            string msg = string.Empty;
            
            if ((cooperationMode != "全職" && cooperationMode != "合約" && cooperationMode != "皆可"))
            {
                msg = "合作狀態須為\"全職\"or\"合約\"or\"皆可\"";
            }

            return msg;
        }

        /// <summary>
        /// 驗證此聯繫狀況是否存在
        /// </summary>
        /// <param name="contactStatus">聯繫狀況</param>
        /// <returns>空值代表沒有錯誤，錯誤會回傳"沒有此聯繫狀況"</returns>
        public string ValidContactStatus(string contactStatus)
        {
            string msg = string.Empty;

            Models Contact_Status = new Models();
            List<string> Contact_Statuslist = Contact_Status.Contact_Status;
            if (!Contact_Statuslist.Contains(contactStatus))
            {
                msg = "沒有\"" + contactStatus + "\"此聯繫狀況";
            }

            return msg;
        }

        /// <summary>
        /// 驗證帳號的狀態
        /// </summary>
        /// <param name="states">傳入該帳號的狀態</param>
        /// <returns>啟用回傳true，停用回傳false</returns>
        public bool ValidMemberStates(string states)
        {
            if (states.Equals("啟用"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 確認輸入的兩次新密碼是否相同
        /// </summary>
        /// <param name="newPassword">新密碼</param>
        /// <param name="checkNewPassword">確認新密碼</param>
        /// <returns>空值代表密碼一致</returns>
        public string ValidNewPassword(string newPassword, string checkNewPassword)
        {
            string msg = string.Empty;
            if (!newPassword.Equals(checkNewPassword))
            {
                msg = "新密碼不一致";
            }

            return msg;
        }

        /// <summary>
        /// 驗證專案經驗是否為空值
        /// </summary>
        /// <param name="projectExperienceList"></param>
        /// <returns></returns>
        public string ValidProjectExperienceData(List<ProjectExperience> projectExperienceList)
        {
            string msg = string.Empty;
            foreach (ProjectExperience projectExperience in projectExperienceList)
            {
                if (projectExperience.Company == string.Empty && projectExperience.Project_Name == string.Empty
                    && projectExperience.OS == string.Empty && projectExperience.Database == string.Empty
                    && projectExperience.Description == string.Empty && projectExperience.Position == string.Empty
                    && projectExperience.Start_End_Date == string.Empty && projectExperience.Language == string.Empty
                    && projectExperience.Tools == string.Empty)
                {
                    msg = "有空的專案經驗";
                }
            }

            return msg;
        }

        /// <summary>
        /// 驗證專案經驗是否重複
        /// </summary>
        /// <param name="data">將多個專案經驗包成DataTable</param>
        /// <returns>空值代表沒有錯誤</returns>
        public string ValidProjectExperienceIsRepeat(DataTable data)
        {
            string msg = string.Empty;
            var repeat = (from row in data.AsEnumerable()
                          group row by new
                          {
                              Company = row.Field<string>("Company"),
                              Project_Name = row.Field<string>("Project_Name")
                          } into g
                          where g.Count() > 1
                          select g.Key).ToList();
            if (repeat.Count > 0)
            {
                msg = "專案資料重複";
            }

            return msg;
        }

        /// <summary>
        /// 驗證聯繫狀況資料的狀態的值是否為"追蹤"or"保留"
        /// </summary>
        /// <param name="states">狀態</param>
        /// <returns>空代表沒有錯誤</returns>
        public string ValidStates(string states)
        {
            string msg = string.Empty;
            if (string.IsNullOrEmpty(states))
            {
                return msg;
            }

            if ((states != "追蹤" && states != "儲存") || string.IsNullOrEmpty(states))
            {
                msg = "合作狀態須為\"追蹤\"or\"儲存\"";
            }

            return msg;
        }

        /// <summary>
        /// 驗證面試基本資料是否正確
        /// </summary>
        /// <param name="vacancies">應徵職缺不可為空值</param>
        /// <param name="name">姓名不可為空值</param>
        /// <param name="married">值為"已婚"or"未婚"</param>
        /// <param name="interviewDate">面試日期不可為空值</param>
        /// <param name="sex">值為"男"or"女"</param>
        /// <param name="mail">值為e-mail格式</param>
        /// <param name="birthday">生日</param>
        /// <param name="cellPhone">值為手機格式</param>
        /// <param name="picture">值為照片路徑</param>
        /// <returns>空值代表沒有錯誤</returns>
        public string ValidInterviewInfoData(string vacancies, string name, string married, string interviewDate, string sex, string mail, string birthday, string cellPhone, string picture)
        {
            string msg = string.Empty;
            string validMsg = string.Empty;
            if (string.IsNullOrEmpty(vacancies))
            {
                msg += "應徵職缺不可為空值\n";
            }

            if (string.IsNullOrEmpty(name))
            {
                msg += "姓名不可為空值\n";
            }

            validMsg = Valid.GetInstance().ValidMarried(married);
            if (validMsg != string.Empty)
            {
                msg += validMsg + "\n";
            }

            validMsg = Valid.GetInstance().ValidDateFormat(interviewDate);
            if (validMsg != string.Empty)
            {
                msg += "面談日期為空值or格式不正確\n";
            }

            validMsg = Valid.GetInstance().ValidSex(sex);
            if (validMsg != string.Empty)
            {
                msg += validMsg + "\n";
            }

            if (!string.IsNullOrEmpty(string.Empty))
            {
                validMsg = Valid.GetInstance().ValidDateFormat(birthday);
                if (validMsg != string.Empty)
                {
                    msg += "生日為格式不正確\n";
                }
            }

            return msg;
        }

        /// <summary>
        /// 驗證是否為本公司之帳號
        /// </summary>
        /// <param name="account">帳號</param>
        /// <returns>回傳空值代表沒有錯誤，若不是本公司之帳號則回傳"帳號須為本公司之帳號"</returns>
        public string ValidIsCompanyMail(string account)
        {
            string msg = string.Empty;
            try
            {
                string[] splitAccount = account.Split('@');
                if (!splitAccount[1].Equals("is-land.com.tw"))
                {
                    msg = "帳號須為本公司之帳號";
                }

                return msg;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                msg = "帳號須為本公司之帳號";
                return msg;
            }
        }

        /// <summary>
        /// 檢查代碼正確性
        /// </summary>
        /// <returns></returns>
        public string ValidCode(List<TextBox> codeTxt)
        {
            string msg = string.Empty;
            ////如果代碼超過兩筆則要檢查不能為空值
            if (codeTxt.Count > 1)
            {
                foreach (TextBox code in codeTxt)
                {
                    if (code.Text == string.Empty)
                    {
                        msg = "代碼不可為空值";
                        return msg;
                    }
                }
            }
            
            ////檢查代碼是否重複
            var isdifference = (from code in codeTxt
                                group code by new
                                {
                                    Code_Id = code.Text
                                } into g
                                where g.Count() > 1
                                select g.Key).ToList();
            if (isdifference.Count > 0)
            {
                msg = "輸入重複的代碼";
                return msg;
            }

            return msg;
        }
        
        /// <summary>
        /// 驗證Excel格式為舊版或新版
        /// </summary>
        /// <param name="dataColumn"></param>
        /// <param name="IsOld"></param>
        /// <returns></returns>
        public bool ValidTalentFormat(DataColumnCollection dataColumn, bool IsOld)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@".\Template\TalentVersionTemplate.xlsx");
            Worksheet sheet = workbook.Worksheets[IsOld ? 0 : 1];
            DataTable dt = sheet.ExportDataTable();

            return ValidTalentDataColumn(dataColumn, dt);
        }

        private bool ValidTalentDataColumn(DataColumnCollection dataColumn,DataTable dataTable)
        {
            for (int i = 0; i < 10; i++)
            {
                if (!dataTable.Columns[i].ToString().Equals(dataColumn[i].ToString()))
                {
                    return false;
                }
            }

            return true;
        }
    }
}
