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
        /// 將指定帳號的狀態更改成啟用或停用
        /// </summary>
        /// <param name="account">欲更改狀態的帳號</param>
        /// <returns>啟用(停用)成功(失敗)</returns>
        public string UpdateMemberStatesByAccount(string account, string states)
        {
            string update = @"update Member set States=@states where account=@account";
            try
            {
                using (SqlCommand cmd = new SqlCommand(update, ScConnection, StTransaction))
                {
                    cmd.Parameters.Add("@account", SqlDbType.VarChar).Value = account;
                    cmd.Parameters.Add("@states", SqlDbType.NVarChar).Value = states;
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();
                    return states + "成功";
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return states + "失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }
        /// <summary>
        /// 根據Table執行修改動作
        /// </summary>
        /// <param name="updateData">要修改的資料</param>
        /// <returns>修改成功或失敗</returns>
        public string UpdateData(DataTable updateData, string tableName)
        {
            string update = this.CombinationUpdateSQL(tableName);
            try
            {
                using (SqlCommand cmd = new SqlCommand(update, ScConnection, StTransaction))
                {
                    foreach (DataRow dr in updateData.Rows)
                    {
                        cmd.Parameters.Clear();
                        this.CombinationUpdateParameters(cmd, dr, tableName);
                        cmd.ExecuteNonQuery();
                    }

                    this.CommitTransaction();
                    return "修改成功";
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "修改失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }
        /// <summary>
        /// 將圖片寫入資料庫
        /// </summary>
        /// <param name="interviewId">面談ID</param>
        /// <param name="path">server端路徑</param>
        /// <returns></returns>
        public string UpdateImageByInterviewId(string interviewId, string path)
        {
            if (!this.ValidInterviewIdIsAppear(interviewId))
            {
                return "此面試基本資料，沒有對應的聯繫資料";
            }

            string update = @"update Interview_Info set Image = @image where Interview_Id = @Interview_Id";
            try
            {
                using (SqlCommand cmd = new SqlCommand(update, ScConnection, StTransaction))
                {
                    cmd.Parameters.Add("@image", SqlDbType.NVarChar).Value = path;
                    cmd.Parameters.Add(@"Interview_Id", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();

                    cmd.Parameters.Clear();
                    cmd.CommandText = @"update Contact_Info set UpdateTime = @updateTime where Contact_Id = 
                                       (select Contact_Id from Interview_Info where Interview_Id = @interviewId)";
                    cmd.Parameters.Add("@updateTime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();

                    this.CommitTransaction();
                }

                return "圖片上傳成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "圖片上傳失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 修改最後編輯時間
        /// </summary>
        /// <param name="contactId"></param>
        /// <param name="cmd"></param>
        private void UpdateEditTime(string contactId, SqlCommand cmd)
        {
            cmd.Parameters.Clear();
            cmd.CommandText = @"update Contact_Info set UpdateTime = @updateTime where Contact_Id = @contactId";
            cmd.Parameters.Add("@updateTime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
            cmd.Parameters.Add("@contactId", SqlDbType.Int).Value = contactId;
            cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 修改聯繫基本資料
        /// </summary>
        /// <param name="editData">聯繫基本資料</param>
        /// <param name="id">聯繫ID</param>
        /// <returns>修改成功或失敗</returns>
        public string UpdateContactSituationInfoData(DataTable editData, string id)
        {
            if (!this.ValidIdIsAppear(id))
            {
                return "修改失敗";
            }

            string update = @"update Contact_Info set Name=@name,Sex=@sex,Mail=@mail,CellPhone=@cellPhone,UpdateTime=@updateTime,
                                                      Cooperation_Mode=@cooperationMode,Status=@status,Place=@place,Skill=@skill,Year=@year
                              where Contact_Id=@id";
            try
            {
                using (SqlCommand cmd = new SqlCommand(update, ScConnection, StTransaction))
                {
                    foreach (DataRow dr in editData.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@name", SqlDbType.NVarChar).Value = dr["Name"].ToString();
                        cmd.Parameters.Add("@sex", SqlDbType.NVarChar).Value = dr["Sex"].ToString();
                        cmd.Parameters.Add("@mail", SqlDbType.VarChar).Value = dr["Mail"].ToString();
                        cmd.Parameters.Add("@cellPhone", SqlDbType.VarChar).Value = dr["CellPhone"].ToString();
                        cmd.Parameters.Add("@updateTime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                        cmd.Parameters.Add("@cooperationMode", SqlDbType.NChar).Value = dr["Cooperation_Mode"].ToString();
                        cmd.Parameters.Add("@status", SqlDbType.NVarChar).Value = dr["Status"].ToString();
                        cmd.Parameters.Add("@place", SqlDbType.NVarChar).Value = dr["Place"].ToString();
                        cmd.Parameters.Add("@skill", SqlDbType.NVarChar).Value = dr["Skill"].ToString();
                        cmd.Parameters.Add("@year", SqlDbType.VarChar).Value = dr["Year"].ToString();
                        cmd.Parameters.Add("@id", SqlDbType.Int).Value = id;
                        cmd.ExecuteNonQuery();
                    }

                    this.CommitTransaction();
                    return "修改成功";
                }
            }
            catch (Exception ex)
            {
                this.RollbackTransaction();
                LogInfo.WriteErrorInfo(ex);
                return "修改失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 修改面試基本資料
        /// </summary>
        /// <param name="editData">面試基本資料</param>
        /// <param name="interviewId">面試ID</param>
        /// <returns></returns>
        public string UpdateInterviewInfoData(DataTable editData, string interviewId, string dbPath, string uiPath)
        {
            if (!this.ValidInterviewIdIsAppear(interviewId))
            {
                return "修改失敗";
            }

            string path = TalentFiles.GetInstance().UpdateImage(dbPath, uiPath, interviewId);
            if (path.Equals("上傳失敗") || path.Equals("不存在的路徑") || path.Equals("檔案刪除失敗"))
            {
                throw new Exception();
            }

            string update = @"update Interview_Info set Vacancies=@vacancies,Interview_Date=@interviewDate,Name=@name,Sex=@sex,Birthday=@birthday
                              ,Married=@married,Mail=@mail,Adress=@adress,CellPhone=@cellPhone,Image=@image,Expertise_Language=@expertiseLanguage
                              ,Expertise_Tools=@expertiseTools,Expertise_Tools_Framwork=@expertiseToolsFramwork,Expertise_Devops=@expertiseDevops
                              ,Expertise_OS=@expertiseOS
                              ,Expertise_BigData=@expertiseBigData,Expertise_DataBase=@expertiseDataBase,Expertise_Certification=@expertiseCertification
                              ,IsStudy=@isStudy,IsService=@IsService,Relatives_Relationship=@relativesRelationship,Relatives_Name=@relativesName
                              ,Care_Work=@careWork,Hope_Salary=@hopeSalary,When_Report=@whenReport,Advantage=@advantage,Disadvantages=@disadvantages
                              ,Hobby=@hobby,Attract_Reason=@attractReason,Future_Goal=@futureGoal,Hope_Supervisor=@hopeSupervisor
                              ,Hope_Promise=@hopePromise,Introduction=@introduction,During_Service=@duringService,Exemption_Reason=@exemptionReason
                              ,Urgent_Contact_Person=@urgentContactPerson,Urgent_Relationship=@urgentRelationship,Urgent_CellPhone=@urgentCellPhone
                              ,Education=@education,Language=@language,Work_Experience=@workExperience
                              where Interview_Id=@interviewId";
            try
            {
                using (SqlCommand cmd = new SqlCommand(update, ScConnection, StTransaction))
                {
                    foreach (DataRow dr in editData.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@education", SqlDbType.NVarChar).Value = dr["Education"].ToString().Trim();
                        cmd.Parameters.Add("@language", SqlDbType.NVarChar).Value = dr["Language"].ToString().Trim();
                        cmd.Parameters.Add("@workExperience", SqlDbType.NVarChar).Value = dr["Work_Experience"].ToString().Trim();
                        cmd.Parameters.Add("@vacancies", SqlDbType.NVarChar).Value = dr["Vacancies"].ToString().Trim();
                        cmd.Parameters.Add("@urgentCellPhone", SqlDbType.VarChar).Value = dr["Urgent_CellPhone"].ToString().Trim();
                        cmd.Parameters.Add("@urgentRelationship", SqlDbType.NVarChar).Value = dr["Urgent_Relationship"].ToString().Trim();
                        cmd.Parameters.Add("@urgentContactPerson", SqlDbType.NVarChar).Value = dr["Urgent_Contact_Person"].ToString().Trim();
                        cmd.Parameters.Add("@exemptionReason", SqlDbType.NVarChar).Value = dr["Exemption_Reason"].ToString().Trim();
                        cmd.Parameters.Add("@duringService", SqlDbType.NVarChar).Value = dr["During_Service"].ToString().Trim();
                        cmd.Parameters.Add("@introduction", SqlDbType.NVarChar).Value = dr["Introduction"].ToString().Trim();
                        cmd.Parameters.Add("@hopePromise", SqlDbType.NVarChar).Value = dr["Hope_Promise"].ToString().Trim();
                        cmd.Parameters.Add("@hopeSupervisor", SqlDbType.NVarChar).Value = dr["Hope_Supervisor"].ToString().Trim();
                        cmd.Parameters.Add("@futureGoal", SqlDbType.NVarChar).Value = dr["Future_Goal"].ToString().Trim();
                        cmd.Parameters.Add("@attractReason", SqlDbType.NVarChar).Value = dr["Attract_Reason"].ToString().Trim();
                        cmd.Parameters.Add("@hobby", SqlDbType.NVarChar).Value = dr["Hobby"].ToString().Trim();
                        cmd.Parameters.Add("@disadvantages", SqlDbType.NVarChar).Value = dr["Disadvantages"].ToString().Trim();
                        cmd.Parameters.Add("@advantage", SqlDbType.NVarChar).Value = dr["Advantage"].ToString().Trim();
                        cmd.Parameters.Add("@whenReport", SqlDbType.NVarChar).Value = dr["When_Report"].ToString().Trim();
                        cmd.Parameters.Add("@hopeSalary", SqlDbType.NVarChar).Value = dr["Hope_salary"].ToString().Trim();
                        cmd.Parameters.Add("@careWork", SqlDbType.NVarChar).Value = dr["Care_Work"].ToString().Trim();
                        cmd.Parameters.Add("@relativesName", SqlDbType.NVarChar).Value = dr["Relatives_Name"].ToString().Trim();
                        cmd.Parameters.Add("@relativesRelationship", SqlDbType.NVarChar).Value = dr["Relatives_Relationship"].ToString().Trim();
                        cmd.Parameters.Add("@isService", SqlDbType.NVarChar).Value = dr["IsService"].ToString().Trim();
                        cmd.Parameters.Add("@isStudy", SqlDbType.NVarChar).Value = dr["IsStudy"].ToString().Trim();
                        cmd.Parameters.Add("@expertiseCertification", SqlDbType.NVarChar).Value = dr["Expertise_Certification"].ToString().Trim();
                        cmd.Parameters.Add("@expertiseDataBase", SqlDbType.NVarChar).Value = dr["Expertise_DataBase"].ToString().Trim();
                        cmd.Parameters.Add("@expertiseBigData", SqlDbType.NVarChar).Value = dr["Expertise_BigData"].ToString().Trim();
                        cmd.Parameters.Add("@expertiseOS", SqlDbType.NVarChar).Value = dr["Expertise_OS"].ToString().Trim();
                        cmd.Parameters.Add("@expertiseDevops", SqlDbType.NVarChar).Value = dr["Expertise_Devops"].ToString().Trim();
                        cmd.Parameters.Add("@expertiseTools", SqlDbType.NVarChar).Value = dr["Expertise_Tools"].ToString().Trim();
                        cmd.Parameters.Add("@expertiseToolsFramwork", SqlDbType.NVarChar).Value = dr["Expertise_Tools_Framwork"].ToString().Trim();
                        cmd.Parameters.Add("@expertiseLanguage", SqlDbType.NVarChar).Value = dr["Expertise_language"].ToString().Trim();
                        cmd.Parameters.Add("@image", SqlDbType.NVarChar).Value = path.Trim();
                        cmd.Parameters.Add("@interviewDate", SqlDbType.Date).Value = dr["Interview_Date"].ToString().Trim();
                        cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                        cmd.Parameters.Add("@name", SqlDbType.NVarChar).Value = dr["Name"].ToString().Trim();
                        cmd.Parameters.Add("@sex", SqlDbType.NVarChar).Value = dr["Sex"].ToString().Trim();
                        cmd.Parameters.Add("@birthday", SqlDbType.Date).Value = Common.GetInstance().ValueIsNullOrEmpty(dr["Birthday"].ToString().Trim());
                        cmd.Parameters.Add("@married", SqlDbType.NVarChar).Value = dr["Married"].ToString().Trim();
                        cmd.Parameters.Add("@mail", SqlDbType.VarChar).Value = dr["Mail"].ToString().Trim();
                        cmd.Parameters.Add("@adress", SqlDbType.NVarChar).Value = dr["Adress"].ToString().Trim();
                        cmd.Parameters.Add("@cellPhone", SqlDbType.NVarChar).Value = dr["CellPhone"].ToString().Trim();
                        cmd.ExecuteNonQuery();
                    }

                    cmd.Parameters.Clear();
                    cmd.CommandText = @"update Contact_Info set UpdateTime = @updateTime where Contact_Id = 
                                       (select Contact_Id from Interview_Info where Interview_Id = @interviewId)";
                    cmd.Parameters.Add("@updateTime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                    cmd.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    cmd.ExecuteNonQuery();

                    this.CommitTransaction();
                    return "修改成功";
                }
            }
            catch (Exception ex)
            {
                this.RollbackTransaction();
                LogInfo.WriteErrorInfo(ex);
                return "修改失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        public string UpdatePasswordByaccount(string account, string newPassword)
        {
            string password = Common.GetInstance().PasswordEncryption(newPassword.ToLower());
            string update = @"update Member set Password=@password where account=@account";
            try
            {
                using (SqlCommand cmd = new SqlCommand(update, ScConnection, StTransaction))
                {
                    cmd.Parameters.Add("@password", SqlDbType.VarChar).Value = password;
                    cmd.Parameters.Add("@account", SqlDbType.VarChar).Value = account;
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();
                    return "修改成功";
                }
            }
            catch (Exception ex)
            {
                this.RollbackTransaction();
                LogInfo.WriteErrorInfo(ex);
                return "修改失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 更換密碼功能
        /// </summary>
        /// <param name="account">帳號</param>
        /// <param name="oldPassword">舊密碼</param>
        /// <param name="newPassword">新密碼</param>
        /// <param name="checkNewPassword">確認新密碼</param>
        /// <returns>是否修改成功</returns>
        public string UpdatePasswordByaccount(string account, string oldPassword, string newPassword, string checkNewPassword)
        {
            if (!this.SignIn(account, oldPassword).Equals(account))
            {
                return "密碼錯誤";
            }

            string msg = TalentValid.GetInstance().ValidNewPassword(newPassword, checkNewPassword);
            if (!msg.Equals(string.Empty))
            {
                return msg;
            }

            string password = Common.GetInstance().PasswordEncryption(newPassword.ToLower());
            string update = @"update Member set Password=@password where account=@account";
            try
            {
                using (SqlCommand cmd = new SqlCommand(update, ScConnection, StTransaction))
                {
                    cmd.Parameters.Add("@password", SqlDbType.VarChar).Value = password;
                    cmd.Parameters.Add("@account", SqlDbType.VarChar).Value = account;
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();
                    return "修改成功";
                }
            }
            catch (Exception ex)
            {
                this.RollbackTransaction();
                LogInfo.WriteErrorInfo(ex);
                return "修改失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }
        
    }
}
