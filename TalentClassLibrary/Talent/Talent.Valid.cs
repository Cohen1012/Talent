using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using ShareClassLibrary;
using System.Windows.Forms;
using System.IO;
using System.Drawing;
using TalentClassLibrary.Model;
using System.Diagnostics;

namespace TalentClassLibrary
{
    /// <summary>
    /// 操作資料庫類別
    /// </summary>
    public partial class Talent : SQL
    {
        /// <summary>
        /// 檢查面談資料ID是否存在
        /// </summary>
        /// <param name="InterviewId"></param>
        /// <returns></returns>
        public bool ValidInterviewIdIsAppear(string InterviewId)
        {
            bool msg = false;
            string select = @"select count(1) from Interview_Info where Interview_Id = @interviewId";
            try
            {
                using (SqlCommand cmd = new SqlCommand(select, ScConnection))
                {
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = InterviewId;
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            if (int.Parse(dr[0].ToString()) > 0)
                            {
                                msg = true;
                            }
                            else
                            {
                                msg = false;
                            }
                        }
                    }
                }
                return msg;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return false;
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 驗證此聯繫ID是否存在
        /// </summary>
        /// <param name="id">欲驗證的聯繫ID</param>
        /// <returns>True or False</returns>
        public bool ValidIdIsAppear(string id)
        {
            bool msg = false;
            string select = @"select count(1) from Contact_Info where Contact_Id = @contactId";
            try
            {
                using (SqlCommand cmd = new SqlCommand(select, ScConnection))
                {
                    cmd.Parameters.Add("@contactId", SqlDbType.Int).Value = id;
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            if (int.Parse(dr[0].ToString()) > 0)
                            {
                                msg = true;
                            }
                            else
                            {
                                msg = false;
                            }
                        }
                    }
                }
                return msg;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return false;
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 驗證從Excel匯入的代碼是否已存在於資料庫
        /// </summary>
        /// <param name="codeList"></param>
        /// <returns></returns>
        public string ValidCodeIsRepeat(List<string> codeList)
        {
            string msg = string.Empty;
            DataTable dt = new DataTable();
            string select = @"select Code_Id from Code where Code_Id in (";

            if (codeList.Count == 0)
            {
                return msg;
            }

            for (int i = 0; i < codeList.Count; i++)
            {
                select += @"@codeId" + i + ",";
            }

            select = select.RemoveEndWithDelimiter(",") + ")";

            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    for (int i = 0; i < codeList.Count; i++)
                    {
                        da.SelectCommand.Parameters.Add("@codeId" + i, SqlDbType.VarChar).Value = codeList[i];
                    }

                    da.Fill(dt);
                }

                if(dt.Rows.Count == 0)
                {
                    return msg;
                }

                foreach(DataRow dr in dt.Rows)
                {
                    msg += dr[0].ToString() + "此代碼已存在\n";
                }

                return msg;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                msg = "資料庫錯誤";
                return msg;
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 驗證欲新增，修改的代碼是否已存在
        /// </summary>
        /// <param name="codeList">代碼陣列</param>
        /// <returns>空值代表沒有錯誤</returns>
        public string ValidCodeIsRepeat(DataTable codeList)
        {
            string msg = string.Empty;
            string select = @"select count(1) from Code where Code_Id = @codeId and Contact_Id != @Contact_Id";
            try
            {
                using (SqlCommand cmd = new SqlCommand(select, ScConnection))
                {
                    foreach (DataRow row in codeList.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@codeId", SqlDbType.VarChar).Value = row["Code_Id"].ToString();
                        cmd.Parameters.Add("@Contact_Id", SqlDbType.VarChar).Value = row["Contact_Id"].ToString();
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            if (dr.Read())
                            {
                                if (int.Parse(dr[0].ToString()) > 0)
                                {
                                    msg += row["Code_Id"].ToString() + "此代碼已存在\n";
                                }
                            }
                        }
                    }

                }
                return msg;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                msg = "發生未預期錯誤";
                return msg;
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 驗證欲新增的帳號是否已存在
        /// </summary>
        /// <param name="account">欲新增的帳號</param>
        /// <returns>空值代表不存在</returns>
        public string ValidInsertMember(string account)
        {
            if (string.IsNullOrEmpty(account))
            {
                return "帳號不可為空值";
            }

            int count = 0;
            account = TalentCommon.GetInstance().AddMailFormat(account);
            string msg = TalentValid.GetInstance().ValidIsCompanyMail(account);
            if (!msg.Equals(string.Empty))
            {
                return msg;
            }

            string select = @"select count(1) from Member where account=@account";
            try
            {
                using (SqlCommand cmd = new SqlCommand(select, ScConnection))
                {
                    cmd.Parameters.Add("@account", SqlDbType.VarChar).Value = account;
                    using (SqlDataReader re = cmd.ExecuteReader())
                    {
                        if (re.Read())
                        {
                            count = int.Parse(re[0].ToString());
                        }

                        if (count > 0)
                        {
                            return "帳號已存在";
                        }
                        else
                        {
                            return string.Empty;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "發生錯誤";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }
    }
}
