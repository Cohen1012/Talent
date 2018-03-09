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
        /// 根據Table執行新增動作
        /// </summary>
        /// <param name="inData">要新增的資料</param>
        /// <param name="id">資料對應的聯繫ID或面試ID</param>
        /// <param name="tableName">要新增至哪個Table</param>
        /// <returns>新增成功或失敗</returns>
        public string InsertData(DataTable inData, string id, string tableName)
        {
            string insert = this.CombinationInsertSQL(tableName);
            try
            {
                using (SqlCommand cmd = new SqlCommand(insert, ScConnection, StTransaction))
                {
                    foreach (DataRow dr in inData.Rows)
                    {
                        cmd.Parameters.Clear();
                        this.CombinationInsertParameters(id, cmd, dr, tableName);
                        cmd.ExecuteNonQuery();
                    }

                    this.CommitTransaction();
                    return "新增成功";
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "新增失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 新增面試基本資料
        /// </summary>
        /// <param name="inData"></param>
        /// <param name="contactId"></param>
        /// <returns></returns>
        public string InsertInterviewInfoData(DataTable inData, string contactId)
        {
            if (!this.ValidIdIsAppear(contactId))
            {
                return "此面試基本資料，沒有對應的聯繫資料";
            }

            string interviewId = string.Empty;
            string insert = @"insert into Interview_Info ([Vacancies],[Interview_Date],[Contact_Id],[Name],[Sex],[Birthday],[Married],[Mail]
                              ,[Adress],[CellPhone],[Expertise_Language],[Expertise_Tools],[Expertise_Tools_Framwork],[Expertise_Devops],[Expertise_OS]
                              ,[Expertise_BigData],[Expertise_DataBase],[Expertise_Certification],[IsStudy],[IsService],[Relatives_Relationship]
                              ,[Relatives_Name],[Care_Work],[Hope_Salary],[When_Report],[Advantage],[Disadvantages],[Hobby],[Attract_Reason]
                              ,[Future_Goal],[Hope_Supervisor],[Hope_Promise],[Introduction],[During_Service]
                              ,[Exemption_Reason],[Urgent_Contact_Person],[Urgent_Relationship],[Urgent_CellPhone],[Education],[Language],[Work_Experience])
                               output inserted.Interview_Id values 
                              (@vacancies,@interviewDate,@contactId,@name,@sex,@birthday,@married,@mail,@adress,@cellPhone,@expertiseLanguage
                              ,@expertiseTools,@expertiseToolsFramwork,@expertiseDevops,@expertiseOS,@expertiseBigData,@expertiseDataBase,@expertiseCertification
                              ,@isStudy,@IsService,@relativesRelationship,@relativesName,@careWork,@hopeSalary,@whenReport,@advantage
                              ,@disadvantages,@hobby,@attractReason,@futureGoal,@hopeSupervisor,@hopePromise,@introduction,@duringService
                              ,@exemptionReason,@urgentContactPerson,@urgentRelationship,@urgentCellPhone,@education,@language,@workExperience)";
            try
            {
                using (SqlCommand cmd = new SqlCommand(insert, ScConnection, StTransaction))
                {
                    foreach (DataRow dr in inData.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@education", SqlDbType.NVarChar).Value = dr["Education"].ToString();
                        cmd.Parameters.Add("@language", SqlDbType.NVarChar).Value = dr["Language"].ToString();
                        cmd.Parameters.Add("@workExperience", SqlDbType.NVarChar).Value = dr["Work_Experience"].ToString();
                        cmd.Parameters.Add("@vacancies", SqlDbType.NVarChar).Value = dr["Vacancies"].ToString();
                        cmd.Parameters.Add("@urgentCellPhone", SqlDbType.VarChar).Value = dr["Urgent_CellPhone"].ToString();
                        cmd.Parameters.Add("@urgentRelationship", SqlDbType.NVarChar).Value = dr["Urgent_Relationship"].ToString();
                        cmd.Parameters.Add("@urgentContactPerson", SqlDbType.NVarChar).Value = dr["Urgent_Contact_Person"].ToString();
                        cmd.Parameters.Add("@exemptionReason", SqlDbType.NVarChar).Value = dr["Exemption_Reason"].ToString();
                        cmd.Parameters.Add("@duringService", SqlDbType.NVarChar).Value = dr["During_Service"].ToString();
                        cmd.Parameters.Add("@introduction", SqlDbType.NVarChar).Value = dr["Introduction"].ToString();
                        cmd.Parameters.Add("@hopePromise", SqlDbType.NVarChar).Value = dr["Hope_Promise"].ToString();
                        cmd.Parameters.Add("@hopeSupervisor", SqlDbType.NVarChar).Value = dr["Hope_Supervisor"].ToString();
                        cmd.Parameters.Add("@futureGoal", SqlDbType.NVarChar).Value = dr["Future_Goal"].ToString();
                        cmd.Parameters.Add("@attractReason", SqlDbType.NVarChar).Value = dr["Attract_Reason"].ToString();
                        cmd.Parameters.Add("@hobby", SqlDbType.NVarChar).Value = dr["Hobby"].ToString();
                        cmd.Parameters.Add("@disadvantages", SqlDbType.NVarChar).Value = dr["Disadvantages"].ToString();
                        cmd.Parameters.Add("@advantage", SqlDbType.NVarChar).Value = dr["Advantage"].ToString();
                        cmd.Parameters.Add("@whenReport", SqlDbType.NVarChar).Value = dr["When_Report"].ToString();
                        cmd.Parameters.Add("@hopeSalary", SqlDbType.NVarChar).Value = dr["Hope_salary"].ToString();
                        cmd.Parameters.Add("@careWork", SqlDbType.NVarChar).Value = dr["Care_Work"].ToString();
                        cmd.Parameters.Add("@relativesName", SqlDbType.NVarChar).Value = dr["Relatives_Name"].ToString();
                        cmd.Parameters.Add("@relativesRelationship", SqlDbType.NVarChar).Value = dr["Relatives_Relationship"].ToString();
                        cmd.Parameters.Add("@isService", SqlDbType.NVarChar).Value = dr["IsService"].ToString();
                        cmd.Parameters.Add("@isStudy", SqlDbType.NVarChar).Value = dr["IsStudy"].ToString();
                        cmd.Parameters.Add("@expertiseCertification", SqlDbType.NVarChar).Value = dr["Expertise_Certification"].ToString();
                        cmd.Parameters.Add("@expertiseDataBase", SqlDbType.NVarChar).Value = dr["Expertise_DataBase"].ToString();
                        cmd.Parameters.Add("@expertiseBigData", SqlDbType.NVarChar).Value = dr["Expertise_BigData"].ToString();
                        cmd.Parameters.Add("@expertiseOS", SqlDbType.NVarChar).Value = dr["Expertise_OS"].ToString();
                        cmd.Parameters.Add("@expertiseDevops", SqlDbType.NVarChar).Value = dr["Expertise_Devops"].ToString();
                        cmd.Parameters.Add("@expertiseTools", SqlDbType.NVarChar).Value = dr["Expertise_Tools"].ToString();
                        cmd.Parameters.Add("@expertiseToolsFramwork", SqlDbType.NVarChar).Value = dr["Expertise_Tools_Framwork"].ToString();
                        cmd.Parameters.Add("@expertiseLanguage", SqlDbType.NVarChar).Value = dr["Expertise_language"].ToString();
                        cmd.Parameters.Add("@interviewDate", SqlDbType.Date).Value = dr["Interview_Date"].ToString();
                        cmd.Parameters.Add("@contactId", SqlDbType.Int).Value = contactId;
                        cmd.Parameters.Add("@name", SqlDbType.NVarChar).Value = dr["Name"].ToString();
                        cmd.Parameters.Add("@sex", SqlDbType.NVarChar).Value = dr["Sex"].ToString();
                        cmd.Parameters.Add("@birthday", SqlDbType.Date).Value = Common.GetInstance().ValueIsNullOrEmpty(dr["Birthday"].ToString());
                        cmd.Parameters.Add("@married", SqlDbType.NVarChar).Value = dr["Married"].ToString();
                        cmd.Parameters.Add("@mail", SqlDbType.VarChar).Value = dr["Mail"].ToString();
                        cmd.Parameters.Add("@adress", SqlDbType.NVarChar).Value = dr["Adress"].ToString();
                        cmd.Parameters.Add("@cellPhone", SqlDbType.VarChar).Value = dr["CellPhone"].ToString();
                        interviewId = cmd.ExecuteScalar().ToString();
                        if (!string.IsNullOrEmpty(dr["Image"].ToString()))
                        {
                            string path = TalentFiles.GetInstance().UpLoadImage(dr["Image"].ToString(), interviewId);
                            if (path.Equals("上傳失敗") || path.Equals("不存在的路徑"))
                            {
                                throw new Exception(path);
                            }

                            cmd.CommandText = @"update Interview_Info set Image = @image where Interview_Id = @Interview_Id";
                            cmd.Parameters.Add("@image", SqlDbType.NVarChar).Value = path;
                            cmd.Parameters.Add("@Interview_Id", SqlDbType.Int).Value = interviewId;
                            cmd.ExecuteNonQuery();
                        }
                    }

                    this.UpdateEditTime(contactId, cmd);
                    this.CommitTransaction();
                }

                return interviewId;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "新增失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 新增聯繫基本資料
        /// </summary>
        /// <param name="inData"></param>
        /// <returns></returns>
        public string InsertContactSituationInfoData(DataTable inData)
        {
            string contact_Id = string.Empty;
            string insert = @"insert into Contact_Info ([Name],[Sex],[Mail],[CellPhone],[UpdateTime],[Cooperation_Mode],[Status],[Place],[Skill],[Year])
                              output inserted.Contact_Id
                              values(@name,@sex,@mail,@cellPhone,@updateTime,@cooperationMode,@status,@place,@skill,@year)";
            try
            {
                using (SqlCommand cmd = new SqlCommand(insert, ScConnection, StTransaction))
                {
                    foreach (DataRow dr in inData.Rows)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@name", SqlDbType.NVarChar).Value = dr["Name"].ToString();
                        cmd.Parameters.Add("@sex", SqlDbType.NVarChar).Value = dr["Sex"].ToString();
                        cmd.Parameters.Add("@mail", SqlDbType.VarChar).Value = dr["Mail"].ToString();
                        cmd.Parameters.Add("@cellPhone", SqlDbType.VarChar).Value = dr["CellPhone"].ToString();
                        cmd.Parameters.Add("@updateTime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                        cmd.Parameters.Add("@cooperationMode", SqlDbType.NVarChar).Value = dr["Cooperation_Mode"].ToString();
                        cmd.Parameters.Add("@status", SqlDbType.NVarChar).Value = dr["Status"].ToString();
                        cmd.Parameters.Add("@place", SqlDbType.NVarChar).Value = dr["Place"].ToString();
                        cmd.Parameters.Add("@skill", SqlDbType.NVarChar).Value = dr["Skill"].ToString();
                        cmd.Parameters.Add("@year", SqlDbType.VarChar).Value = dr["Year"].ToString();
                        contact_Id = cmd.ExecuteScalar().ToString();
                    }

                    this.CommitTransaction();
                    return contact_Id;
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "新增失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 新增使用者帳號
        /// </summary>
        /// <param name="account">欲新增的帳號</param>
        /// <returns>新增成功或失敗</returns>
        public string InsertMember(string account)
        {
            try
            {
                string msg = this.ValidInsertMember(account);
                if (!msg.Equals(string.Empty))
                {
                    return msg;
                }

                account =  TalentCommon.GetInstance().AddMailFormat(account);
                string[] splitAccount = account.Split('@');
                string password = Common.GetInstance().PasswordEncryption(splitAccount[0].ToLower());
                string insert = @"insert into Member (Account,Password,States) values (@account,@password,N'啟用')";
                using (SqlCommand cmd = new SqlCommand(insert, ScConnection, StTransaction))
                {
                    cmd.Parameters.Add("@account", SqlDbType.VarChar).Value = account;
                    cmd.Parameters.Add("@password", SqlDbType.VarChar).Value = password;
                    cmd.ExecuteNonQuery();
                    this.CommitTransaction();
                    return "新增成功";
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "新增失敗";
            }
        }
    }
}
