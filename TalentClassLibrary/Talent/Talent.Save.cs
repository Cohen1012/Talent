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
        /// 將附加檔案資訊寫入資料庫
        /// </summary>
        /// <param name="files"></param>
        /// <param name="interviewId"></param>
        /// <param name="attachedFileMode"></param>
        /// <returns></returns>
        public string SaveAttachedFilesToDB(List<string> files, string interviewId, string attachedFileMode)
        {
            string sqlStr = string.Empty;
            List<string> path = TalentFiles.GetInstance().SaveAttachedFiles(files, interviewId, attachedFileMode);
            if (path.Contains("上傳失敗"))
            {
                return "檔案上傳失敗";
            }

            try
            {
                using (SqlCommand cmd = new SqlCommand(sqlStr, ScConnection, StTransaction))
                {
                    cmd.CommandText = @"delete from Files where Interview_Id = @Interview_Id and Belong = @Belong";
                    cmd.Parameters.Add("@Interview_Id", SqlDbType.Int).Value = interviewId;
                    cmd.Parameters.Add("@Belong", SqlDbType.NVarChar).Value = attachedFileMode;
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = @"insert into Files (Interview_Id,File_Path,Belong) values (@Interview_Id,@File_Path,@Belong)";
                    foreach (string file in path)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@Interview_Id", SqlDbType.Int).Value = interviewId;
                        cmd.Parameters.Add("@File_Path", SqlDbType.NVarChar).Value = file;
                        cmd.Parameters.Add("@Belong", SqlDbType.NVarChar).Value = attachedFileMode;
                        cmd.ExecuteNonQuery();
                    }

                    cmd.Parameters.Clear();
                    cmd.CommandText = @"update Contact_Info set UpdateTime = @updateTime where Contact_Id = 
                                       (select Contact_Id from Interview_Info where Interview_Id = @interviewId)";
                    cmd.Parameters.Add("@updateTime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();

                    this.CommitTransaction();
                    return "檔案上傳成功";
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "檔案上傳失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }
        
        /// <summary>
        /// 儲存面談結果資料
        /// </summary>
        /// <param name="saveData">面談結果資料</param>
        /// <param name="interviewId">對應的面試ID</param>
        /// <returns>儲存成功或失敗</returns>
        public string SaveInterviewResult(DataSet saveData, string interviewId)
        {
            if (!this.ValidInterviewIdIsAppear(interviewId))
            {
                return "沒有對應的面試基本資料";
            }

            string sqlStr = string.Empty;
            try
            {
                using (SqlCommand cmd = new SqlCommand(sqlStr, ScConnection, StTransaction))
                {
                    cmd.CommandText = @"update Interview_Info set Appointment=@Appointment,Results_Remark=@Results_Remark where Interview_Id=@Interview_Id";
                    foreach (DataRow dr in saveData.Tables[1].Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(@"Appointment", SqlDbType.NVarChar).Value = dr["Appointment"].ToString();
                        cmd.Parameters.Add(@"Results_Remark", SqlDbType.NVarChar).Value = dr["Results_Remark"].ToString();
                        cmd.Parameters.Add(@"Interview_Id", SqlDbType.Int).Value = interviewId;
                        cmd.ExecuteNonQuery();
                    }

                    cmd.Parameters.Clear();
                    cmd.CommandText = @"delete from [Interview_Comments] where Interview_Id=@interviewId";
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = @"insert into [Interview_Comments] ([Interviewer],[Result],[Interview_Id])
                                        values (@Interviewer,@Result,@Interview_Id)";
                    foreach (DataRow dr in saveData.Tables[0].Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(@"Interviewer", SqlDbType.NVarChar).Value = dr["Interviewer"].ToString();
                        cmd.Parameters.Add(@"Result", SqlDbType.NVarChar).Value = dr["Result"].ToString();
                        cmd.Parameters.Add(@"Interview_Id", SqlDbType.Int).Value = interviewId;
                        cmd.ExecuteNonQuery();
                    }

                    cmd.Parameters.Clear();
                    cmd.CommandText = @"update Contact_Info set UpdateTime = @updateTime where Contact_Id = 
                                       (select Contact_Id from Interview_Info where Interview_Id = @interviewId)";
                    cmd.Parameters.Add("@updateTime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();
                }

                return "儲存成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "儲存失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 儲存聯繫狀況資料
        /// </summary>
        /// <param name="saveData">聯繫狀況資料</param>
        /// <param name="contactId">聯繫ID</param>
        /// <returns>儲存成功或失敗</returns>
        public string SaveContactSituation(DataTable saveData, string contactId)
        {
            if (!this.ValidIdIsAppear(contactId))
            {
                return "此聯繫狀況資料，沒有對應的聯繫資料";

            }

            string sqlStr = string.Empty;
            try
            {
                using (SqlCommand cmd = new SqlCommand(sqlStr, ScConnection, StTransaction))
                {
                    cmd.CommandText = @"delete from Contact_Situation where Contact_Id=@Contact_Id";
                    cmd.Parameters.Add("@Contact_Id", SqlDbType.Int).Value = contactId;
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = @"insert into Contact_Situation (Contact_Status,Remarks,Contact_Date,Contact_Id)
                                        values (@Contact_Status,@Remarks,@Contact_Date,@Contact_Id)";
                    foreach (DataRow dr in saveData.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(@"Contact_Status", SqlDbType.NVarChar).Value = dr["Contact_Status"].ToString();
                        cmd.Parameters.Add(@"Remarks", SqlDbType.NVarChar).Value = dr["Remarks"].ToString();
                        cmd.Parameters.Add(@"Contact_Date", SqlDbType.DateTime).Value = dr["Contact_Date"].ToString();
                        cmd.Parameters.Add(@"Contact_Id", SqlDbType.Int).Value = contactId;
                        cmd.ExecuteNonQuery();
                    }

                    this.CommitTransaction();
                }

                return "儲存成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "儲存失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 儲存代碼資料
        /// </summary>
        /// <param name="saveData">代碼資料</param>
        /// <param name="contactId">聯繫ID</param>
        /// <returns>儲存成功或失敗</returns>
        public string SaveCode(DataTable saveData, string contactId)
        {
            if (!this.ValidIdIsAppear(contactId))
            {
                return "此代碼，沒有對應的聯繫資料";
            }

            string sqlStr = string.Empty;
            try
            {
                using (SqlCommand cmd = new SqlCommand(sqlStr, ScConnection, StTransaction))
                {
                    cmd.CommandText = @"delete from Code where Contact_Id=@Contact_Id";
                    cmd.Parameters.Add("@Contact_Id", SqlDbType.Int).Value = contactId;
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = @"insert into Code (Contact_Id,Code_Id)
                                        values (@Contact_Id,@Code_Id)";
                    foreach (DataRow dr in saveData.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(@"Code_Id", SqlDbType.VarChar).Value = dr["Code_Id"].ToString();
                        cmd.Parameters.Add(@"Contact_Id", SqlDbType.Int).Value = contactId;
                        cmd.ExecuteNonQuery();
                    }

                    this.CommitTransaction();
                }

                return "儲存成功";
            }
            catch (Exception ex)
            {
                this.RollbackTransaction();
                LogInfo.WriteErrorInfo(ex);
                return "儲存失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 儲存專案經驗資料
        /// </summary>
        /// <param name="saveData">專案經驗資料</param>
        /// <param name="interviewId">對應的面試ID</param>
        /// <returns>儲存成功或失敗</returns>
        public string SaveProjectExperience(DataTable saveData, string interviewId)
        {
            if (!this.ValidInterviewIdIsAppear(interviewId))
            {
                return "沒有對應的面試基本資料";
            }

            string sqlStr = string.Empty;
            try
            {
                using (SqlCommand cmd = new SqlCommand(sqlStr, ScConnection, StTransaction))
                {
                    cmd.CommandText = @"delete from [Project_Experience] where Interview_Id=@interviewId";
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = @"insert into [Project_Experience] ([Company],[Project_Name],[OS],[Database],[Position],[Language],[Tools],[Description],[Start_End_Date],[Interview_Id])
                                        values (@Company,@Project_Name,@OS,@Database,@Position,@Language,@Tools,@Description,@Start_End_Date,@Interview_Id)";
                    foreach (DataRow dr in saveData.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(@"Company", SqlDbType.NVarChar).Value = dr["Company"].ToString();
                        cmd.Parameters.Add(@"Project_Name", SqlDbType.NVarChar).Value = dr["Project_Name"].ToString();
                        cmd.Parameters.Add(@"OS", SqlDbType.NVarChar).Value = dr["OS"].ToString();
                        cmd.Parameters.Add(@"Database", SqlDbType.NVarChar).Value = dr["Database"].ToString();
                        cmd.Parameters.Add(@"Position", SqlDbType.NVarChar).Value = dr["Position"].ToString();
                        cmd.Parameters.Add(@"Language", SqlDbType.NVarChar).Value = dr["Language"].ToString();
                        cmd.Parameters.Add(@"Tools", SqlDbType.NVarChar).Value = dr["Tools"].ToString();
                        cmd.Parameters.Add(@"Description", SqlDbType.NVarChar).Value = dr["Description"].ToString();
                        cmd.Parameters.Add(@"Start_End_Date", SqlDbType.NVarChar).Value = dr["Start_End_Date"].ToString();
                        cmd.Parameters.Add(@"Interview_Id", SqlDbType.Int).Value = interviewId;
                        cmd.ExecuteNonQuery();
                    }

                    cmd.Parameters.Clear();
                    cmd.CommandText = @"update Contact_Info set UpdateTime = @updateTime where Contact_Id = 
                                       (select Contact_Id from Interview_Info where Interview_Id = @interviewId)";
                    cmd.Parameters.Add("@updateTime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();
                }

                return "儲存成功";
            }
            catch (Exception ex)
            {
                this.RollbackTransaction();
                LogInfo.WriteErrorInfo(ex);
                return "儲存失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 儲存學歷資料
        /// </summary>
        /// <param name="saveData">學歷資料</param>
        /// <param name="interviewId">對應的面試ID</param>
        /// <returns>儲存成功或失敗</returns>
        public string SaveEducation(DataTable saveData, string interviewId)
        {
            if (!this.ValidInterviewIdIsAppear(interviewId))
            {
                return "沒有對應的面試基本資料";
            }

            string sqlStr = string.Empty;
            try
            {
                using (SqlCommand cmd = new SqlCommand(sqlStr, ScConnection, StTransaction))
                {
                    cmd.CommandText = @"delete from Education where Interview_Id=@interviewId";
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = @"insert into Education ([School],[Department],[Start_End_Date],[Is_Graduation],[Remark],[Interview_Id])
                                        values (@School,@Department,@Start_End_Date,@Is_Graduation,@Remark,@Interview_Id)";
                    foreach (DataRow dr in saveData.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(@"School", SqlDbType.NVarChar).Value = dr["School"].ToString();
                        cmd.Parameters.Add(@"Department", SqlDbType.NVarChar).Value = dr["Department"].ToString();
                        cmd.Parameters.Add(@"Start_End_Date", SqlDbType.NVarChar).Value = dr["Start_End_Date"].ToString();
                        cmd.Parameters.Add(@"Is_Graduation", SqlDbType.NVarChar).Value = dr["Is_Graduation"].ToString();
                        cmd.Parameters.Add(@"Remark", SqlDbType.NVarChar).Value = dr["Remark"].ToString();
                        cmd.Parameters.Add(@"Interview_Id", SqlDbType.Int).Value = interviewId;
                        cmd.ExecuteNonQuery();
                    }

                    this.CommitTransaction();
                }

                return "儲存成功";
            }
            catch (Exception ex)
            {
                this.RollbackTransaction();
                LogInfo.WriteErrorInfo(ex);
                return "儲存失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 儲存工作經歷資料
        /// </summary>
        /// <param name="saveData">工作經歷資料</param>
        /// <param name="interviewId">對應的面試ID</param>
        /// <returns>儲存成功或失敗</returns>
        public string SaveWorkExperience(DataTable saveData, string interviewId)
        {
            if (!this.ValidInterviewIdIsAppear(interviewId))
            {
                return "沒有對應的面試基本資料";
            }

            string sqlStr = string.Empty;
            try
            {
                using (SqlCommand cmd = new SqlCommand(sqlStr, ScConnection, StTransaction))
                {
                    cmd.CommandText = @"delete from Work_Experience where Interview_Id=@interviewId";
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = @"insert into Work_Experience ([Institution_name],[Position],[Start_End_Date],[Start_Salary],[End_Salary],[Leaving_Reason],[Interview_Id])
                                        values (@Institution_name,@Position,@Start_End_Date,@Start_Salary,@End_Salary,@Leaving_Reason,@Interview_Id)";
                    foreach (DataRow dr in saveData.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(@"Institution_name", SqlDbType.NVarChar).Value = dr["Institution_name"].ToString();
                        cmd.Parameters.Add(@"Position", SqlDbType.NVarChar).Value = dr["Position"].ToString();
                        cmd.Parameters.Add(@"Start_End_Date", SqlDbType.NVarChar).Value = dr["Start_End_Date"].ToString();
                        cmd.Parameters.Add(@"Start_Salary", SqlDbType.NVarChar).Value = dr["Start_Salary"].ToString();
                        cmd.Parameters.Add(@"End_Salary", SqlDbType.NVarChar).Value = dr["End_Salary"].ToString();
                        cmd.Parameters.Add(@"Leaving_Reason", SqlDbType.NVarChar).Value = dr["Leaving_Reason"].ToString();
                        cmd.Parameters.Add(@"Interview_Id", SqlDbType.Int).Value = interviewId;
                        cmd.ExecuteNonQuery();
                    }

                    this.CommitTransaction();
                }

                return "儲存成功";
            }
            catch (Exception ex)
            {
                this.RollbackTransaction();
                LogInfo.WriteErrorInfo(ex);
                return "儲存失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }
    }
}
