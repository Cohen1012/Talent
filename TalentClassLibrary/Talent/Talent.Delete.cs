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
        /// 刪除指定的面談資料
        /// </summary>
        /// <param name="interviewId">面談ID</param>
        /// <returns></returns>
        public string DelInterviewDataByInterviewId(string interviewId)
        {
            if (!this.ValidInterviewIdIsAppear(interviewId))
            {
                return "此面談資料不存在";
            }

            if (this.DelImageByInterviewId(interviewId) == "圖片刪除失敗")
            {

                return "圖片刪除失敗";
            }

            if (TalentFiles.GetInstance().DelFilesByInterviewId(interviewId) == "附加檔案刪除失敗")
            {
                return "附加檔案刪除失敗";
            }

            string sqlStr = string.Empty;
            try
            {
                using (SqlCommand cmd = new SqlCommand(sqlStr, ScConnection, StTransaction))
                {
                    ////更新最後修改時間
                    cmd.CommandText = @"update Contact_Info set UpdateTime = @updateTime where Contact_Id = 
                                       (select Contact_Id from Interview_Info where Interview_Id = @interviewId)";
                    cmd.Parameters.Add("@updateTime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();

                    ////專案經驗資料
                    cmd.CommandText = @"delete from Project_Experience where Interview_Id = @id";
                    cmd.Parameters.Add("@id", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();
                    ////面談評語資料
                    cmd.CommandText = @"delete from Interview_Comments where Interview_Id = @id";
                    cmd.ExecuteNonQuery();
                    ////附加檔案
                    cmd.CommandText = @"delete from Files where Interview_Id = @id";
                    cmd.ExecuteNonQuery();
                    ////面談基本資料
                    cmd.CommandText = @"delete from Interview_Info where Interview_Id = @id";
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();
                }

                return "刪除成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "資料庫發生錯誤";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 根據聯繫ID刪除跟其有關的所有資料
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public string DelTalentById(string id)
        {
            if (this.DelImageByContactId(id) == "圖片刪除失敗")
            {
                return "圖片刪除失敗";
            }

            if (this.DelFilesByContactId(id) == "附加檔案刪除失敗")
            {
                return "附加檔案刪除失敗";
            }

            try
            {
                string del = @"delete from Code where Contact_Id = @id"; ////代碼資料
                using (SqlCommand cmd = new SqlCommand(del, ScConnection, StTransaction))
                {
                    cmd.Parameters.Add("@id", SqlDbType.Int).Value = id;
                    cmd.ExecuteNonQuery();
                    ////聯繫基本資料
                    cmd.CommandText = @"delete from Contact_Info where Contact_Id = @id";
                    cmd.ExecuteNonQuery();
                    ////聯繫狀況資料
                    cmd.CommandText = @"delete from Contact_Situation where Contact_Id = @id";
                    cmd.ExecuteNonQuery();
                    ////專案經驗資料
                    cmd.CommandText = @"delete from Project_Experience where Interview_Id in (select Interview_Id from Interview_Info where Contact_Id = @id)";
                    cmd.ExecuteNonQuery();
                    ////面談評語資料
                    cmd.CommandText = @"delete from Interview_Comments where Interview_Id in (select Interview_Id from Interview_Info where Contact_Id = @id)";
                    cmd.ExecuteNonQuery();
                    ////附加檔案
                    cmd.CommandText = @"delete from Files where Interview_Id in (select Interview_Id from Interview_Info where Contact_Id = @id)";
                    cmd.ExecuteNonQuery();
                    ////面談基本資料
                    cmd.CommandText = @"delete from Interview_Info where Contact_Id = @id";
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();

                }
                return "刪除成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "資料庫發生錯誤";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 刪除對應聯繫ID的代碼
        /// </summary>
        /// <param name="id">聯繫ID</param>
        private void DelCodeById(string id)
        {
            string del = @"delete from Code where Contact_Id = @id";
            try
            {
                using (SqlCommand cmd = new SqlCommand(del, ScConnection, StTransaction))
                {
                    cmd.Parameters.Add("@id", SqlDbType.Int).Value = id;
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 根據Table執行刪除動作
        /// </summary>
        /// <param name="DelData">要刪除的資料</param>
        /// <param name="tableName">要刪除的資料</param>
        /// <returns></returns>
        public string DelData(DataTable DelData, string tableName)
        {
            string del = this.CombinationDelSQL(tableName);
            try
            {
                using (SqlCommand cmd = new SqlCommand(del, ScConnection, StTransaction))
                {
                    foreach (DataRow dr in DelData.Rows)
                    {
                        cmd.Parameters.Clear();
                        this.CombinationDelParameters(cmd, dr, tableName);
                        cmd.ExecuteNonQuery();
                    }

                    this.CommitTransaction();
                    return "刪除成功";
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "刪除失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 根據聯繫ID刪除附加檔案
        /// </summary>
        /// <param name="contactId"></param>
        /// <returns></returns>
        public string DelFilesByContactId(string contactId)
        {
            DataTable dt = new DataTable();
            string select = @"select Interview_Id from Interview_Info where Contact_Id = @id";
            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    da.SelectCommand.Parameters.Add("@id", SqlDbType.Int).Value = contactId;
                    da.Fill(dt);
                }

                foreach (DataRow dr in dt.Rows)
                {
                    string serverDir = @".\files\" + dr[0].ToString().Trim();
                    if (Directory.Exists(serverDir))
                    {
                        Directory.Delete(serverDir, true);
                    }
                }

                return "附加檔案刪除成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "附加檔案刪除失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 刪除圖片檔 聯繫ID
        /// </summary>
        /// <param name="contactId"></param>
        /// <returns></returns>
        public string DelImageByContactId(string contactId)
        {
            string select = @" select Image from Interview_Info where Contact_Id = @id";
            try
            {
                using (SqlCommand cmd = new SqlCommand(select, ScConnection))
                {
                    cmd.Parameters.Add("@id", SqlDbType.Int).Value = contactId;
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            if (File.Exists(dr[0].ToString()))
                            {
                                File.Delete(dr[0].ToString());
                            }
                        }
                    }
                }

                return "圖片刪除成功";
            }
            catch (IOException ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "圖片刪除失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 刪除圖片檔 by面談ID
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public string DelImageByInterviewId(string interviewId)
        {
            string select = @"select Image from Interview_Info where Interview_Id = @Interview_Id";
            try
            {
                using (SqlCommand cmd = new SqlCommand(select, ScConnection))
                {
                    cmd.Parameters.Add("@Interview_Id", SqlDbType.Int).Value = interviewId;
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            if (File.Exists(dr[0].ToString()))
                            {
                                File.Delete(dr[0].ToString());
                            }
                        }
                    }
                }

                return "圖片刪除成功";
            }
            catch (IOException ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "圖片刪除失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 刪除指定帳號
        /// </summary>
        /// <param name="account">欲刪除的帳號</param>
        /// <returns>刪除成功或失敗</returns>
        public string DelMemberByAccount(string account)
        {
            string delete = @"delete from Member where account=@account";
            try
            {
                using (SqlCommand cmd = new SqlCommand(delete, ScConnection, StTransaction))
                {
                    cmd.Parameters.Add("@account", SqlDbType.VarChar).Value = account;
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();
                    return "刪除成功";
                }
            }
            catch (Exception ex)
            {
                this.RollbackTransaction();
                LogInfo.WriteErrorInfo(ex);
                return "刪除失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }
    }
}
