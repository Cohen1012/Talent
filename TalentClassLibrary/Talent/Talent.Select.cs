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
        /// 根據地點，技能，合作模式，聯繫狀態，最後編輯時間，查詢符合的Id
        /// </summary>
        /// <param name="places">地點，多筆請用","隔開</param>
        /// <param name="expertises">技能，多筆請用","隔開</param>
        /// <param name="cooperationMode">合作模式</param>
        /// <param name="states">聯繫狀態</param>
        /// <param name="startEditDate">起始日期，日期格式</param>
        /// <param name="endEditDate">結束日期，日期格式</param>
        /// <returns>回傳符合條件的Id</returns>
        public List<string> SelectIdByContact(string places, string expertises, string cooperationMode, string states, string startEditDate, string endEditDate)
        {
            ErrorMessage = string.Empty;
            DataTable dt = new DataTable();
            List<string> idList = new List<string>();
            try
            {
                if (Valid.GetInstance().ValidDateRange(startEditDate, endEditDate) != string.Empty)
                {
                    MessageBox.Show("最後編輯日之日期格式或者是日期區間不正確", "錯誤訊息");
                    return new List<string>();
                }

                string[] place = places.Split(',');
                string[] expertise = expertises.Split(',');
                string select = @"select Contact_Id from Contact_Info where ISNULL(Status,'NA') = ISNULL(ISNULL(@status,Status),'NA') and
                                                                            UpdateTime >= ISNULL(@startEditDate, UpdateTime) and
                                                                            UpdateTime <= ISNULL(@endEditDate, UpdateTime) and
                                                                            ISNULL(Cooperation_Mode,'NA') = ISNULL(ISNULL(@CooperationMode,Cooperation_Mode),'NA')";
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    ////如果合作模式為"全職"or"合約"，則值為"皆可"也要被查詢出來
                    if (cooperationMode.Equals("全職") || cooperationMode.Equals("合約"))
                    {
                        da.SelectCommand.CommandText += @" or ISNULL(Cooperation_Mode,'NA') = ISNULL(ISNULL(@CooperationMode1,Cooperation_Mode),'NA')";
                        da.SelectCommand.Parameters.Add("@CooperationMode1", SqlDbType.NChar).Value = Common.GetInstance().ValueIsNullOrEmpty("皆可");
                    }
                    ////多筆地點
                    for (int i = 0; i < place.Length; i++)
                    {
                        if (i == 0)
                        {
                            da.SelectCommand.CommandText += @" and ISNULL(Place,'NA') LIKE ISNULL(ISNULL(@place" + (i + 1) + ", Place),'NA')";
                            da.SelectCommand.Parameters.Add("@place" + (i + 1), SqlDbType.NVarChar).Value = Common.GetInstance().ValueIsNullOrEmpty("%" + place[i] + "%");
                        }
                        else
                        {
                            da.SelectCommand.CommandText += @" or ISNULL(Place,'NA') LIKE ISNULL(ISNULL(@place" + (i + 1) + ", Place),'NA')";
                            da.SelectCommand.Parameters.Add("@place" + (i + 1), SqlDbType.NVarChar).Value = Common.GetInstance().ValueIsNullOrEmpty("%" + place[i] + "%");
                        }
                    }
                    ////多筆技能
                    for (int i = 0; i < expertise.Length; i++)
                    {
                        if (i == 0)
                        {
                            da.SelectCommand.CommandText += @" and ISNULL(Skill,'NA') Like ISNULL(ISNULL(@skill" + (i + 1) + ", Skill),'NA')";
                            da.SelectCommand.Parameters.Add("@skill" + (i + 1), SqlDbType.NVarChar).Value = Common.GetInstance().ValueIsNullOrEmpty("%" + expertise[i] + "%");
                        }
                        else
                        {
                            da.SelectCommand.CommandText += @" or ISNULL(Skill,'NA') Like ISNULL(ISNULL(@skill" + (i + 1) + ", Skill),'NA')";
                            da.SelectCommand.Parameters.Add("@skill" + (i + 1), SqlDbType.NVarChar).Value = Common.GetInstance().ValueIsNullOrEmpty("%" + expertise[i] + "%");
                        }
                    }

                    da.SelectCommand.Parameters.Add("@CooperationMode", SqlDbType.NChar).Value = TalentCommon.GetInstance().ValueIsAny(cooperationMode);
                    da.SelectCommand.Parameters.Add("@status", SqlDbType.NVarChar).Value = TalentCommon.GetInstance().ValueIsAny(states);
                    da.SelectCommand.Parameters.Add("@startEditDate", SqlDbType.DateTime).Value = Common.GetInstance().ValueIsNullOrEmpty(startEditDate);
                    da.SelectCommand.Parameters.Add("@endEditDate", SqlDbType.DateTime).Value = Common.GetInstance().ValueIsNullOrEmpty(endEditDate);


                    da.Fill(dt);
                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            idList.Add(dr[0].ToString());
                        }
                    }

                    return idList;
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return new List<string>();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        ///根據關鍵字查詢符合資料的ID
        /// </summary>
        /// <param name="keyWords">關鍵字，如有多個關鍵字請用","隔開</param>
        /// <returns>符合關鍵字的ID</returns>
        public List<string> SelectIdByKeyWord(string keyWords)
        {
            ErrorMessage = string.Empty;
            DataTable dt = new DataTable();
            List<string> idList = new List<string>();
            try
            {
                string select = string.Empty;
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    ////關鍵字為空則撈出所有ID
                    if (string.IsNullOrEmpty(keyWords))
                    {
                        da.SelectCommand.CommandText += @"(select Contact_Id from Contact_Info)";
                        da.Fill(dt);
                        if (dt.Rows.Count > 0)
                        {
                            foreach (DataRow dr in dt.Rows)
                            {
                                idList.Add(dr[0].ToString());
                            }
                        }

                        return idList;
                    }

                    string[] keyWord = keyWords.Split(',');
                    foreach (string word in keyWord)
                    {
                        if (string.IsNullOrEmpty(word))
                        {
                            continue;
                        }

                        da.SelectCommand.Parameters.Clear();
                        da.SelectCommand.CommandType = CommandType.StoredProcedure;
                        da.SelectCommand.CommandText = "KeyWordSearch";
                        da.SelectCommand.Parameters.Add("@KeyWord", SqlDbType.NVarChar).Value = word;
                        Common.GetInstance().ClearDataTable(dt);
                        da.Fill(dt);
                        if (dt.Rows.Count > 0)
                        {
                            idList.AddRange(this.SelectIdbyResult(dt, da));
                        }
                    }

                    return idList.Distinct().ToList();
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return new List<string>();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 根據面談條件查詢符合的聯繫ID
        /// </summary>
        /// <param name="isInterview">是否已面談，值為"已面談","未面談","不限"</param>
        /// <param name="interviewResult">面談結果，值為"錄用","不錄用","暫保留","不限"</param>
        /// <param name="startInterviewDate">起始日期，日期格式</param>
        /// <param name="endInterviewDate">結束日期，日期格式</param>
        /// <returns>回傳符合條件的聯繫ID</returns>
        public List<string> SelectIdByInterviewFilter(string isInterview, string interviewResult, string startInterviewDate, string endInterviewDate)
        {
            ErrorMessage = string.Empty;
            DataTable dt = new DataTable();
            List<string> idList = new List<string>();
            string select = string.Empty;
            try
            {
                if (Valid.GetInstance().ValidDateRange(startInterviewDate, endInterviewDate) != string.Empty)
                {
                    MessageBox.Show("面試日期之日期格式或者是日期區間不正確", "錯誤訊息");
                    return new List<string>();
                }

                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    ////是否已面談
                    if (isInterview == "已面談")
                    {
                        da.SelectCommand.CommandText += @"(select Contact_Id from Interview_Info where Appointment is not null and Appointment !='') INTERSECT ";
                    }
                    else if (isInterview == "未面談")
                    {
                        da.SelectCommand.CommandText += @"(select Contact_Id from Contact_Info
                                                          EXCEPT
                                                          select Contact_Id from Interview_Info where Appointment is not null and Appointment !='') INTERSECT ";
                    }
                    ////面談結果
                    if (interviewResult != "不限")
                    {
                        da.SelectCommand.CommandText += @"(select Contact_Id from Interview_Info where Appointment = ISNULL(@interviewResult, Appointment)) INTERSECT ";
                        da.SelectCommand.Parameters.Add("@interviewResult", SqlDbType.NVarChar).Value = Common.GetInstance().ValueIsNullOrEmpty(interviewResult);
                    }
                    ////面談日期
                    if (!string.IsNullOrEmpty(startInterviewDate) || !string.IsNullOrEmpty(endInterviewDate))
                    {
                        da.SelectCommand.CommandText += @"(select Contact_Id from Interview_Info where Interview_Date <= ISNULL(@endInterviewDate, Interview_Date) AND Interview_Date >= ISNULL(@startInterviewDate, Interview_Date)) INTERSECT ";
                        da.SelectCommand.Parameters.Add("@endInterviewDate", SqlDbType.DateTime).Value = Common.GetInstance().ValueIsNullOrEmpty(endInterviewDate);
                        da.SelectCommand.Parameters.Add("@startInterviewDate", SqlDbType.DateTime).Value = Common.GetInstance().ValueIsNullOrEmpty(startInterviewDate);
                    }
                    ////所有ID
                    da.SelectCommand.CommandText += @"(select Contact_Id from Contact_Info)";
                    da.Fill(dt);
                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            idList.Add(dr[0].ToString());
                        }
                    }

                    return idList;
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return new List<string>();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 根據條件查詢符合的ID
        /// </summary>
        /// <param name="keyWords">關鍵字，如有多個關鍵字請用","隔開</param>
        /// <param name="places">地點，多筆請用","隔開</param>
        /// <param name="expertises">技能，多筆請用","隔開</param>
        /// <param name="cooperationMode">合作模式</param>
        /// <param name="states">聯繫狀態</param>
        /// <param name="startEditDate">起始日期，日期格式</param>
        /// <param name="endEditDate">結束日期，日期格式</param>
        /// <param name="isInterview">是否已面談，值為"已面談","未面談","不限"</param>
        /// <param name="interviewResult">面談結果，值為"錄用","不錄用","暫保留","不限"</param>
        /// <param name="startInterviewDate">起始日期，日期格式</param>
        /// <param name="endInterviewDate">結束日期，日期格式</param>
        /// <returns></returns>
        public List<string> SelectIdByFilter(string keyWords, string places, string expertises, string cooperationMode, string states, string startEditDate, string endEditDate, string isInterview, string interviewResult, string startInterviewDate, string endInterviewDate)
        {
            List<string> idListBykeywords = this.SelectIdByKeyWord(keyWords);
            List<string> idListByContact = this.SelectIdByContact(places, expertises, cooperationMode, states, startEditDate, endEditDate);
            List<string> idListByInterviewFilter = this.SelectIdByInterviewFilter(isInterview, interviewResult, startInterviewDate, endInterviewDate);
            List<string> idList = new List<string>();
            idList = idListBykeywords.Intersect(idListByContact).ToList();
            idList = idList.Intersect(idListByInterviewFilter).ToList();
            return idList;

        }

        ///透過預存程序會找出符合值得所在Table and 欄位
        /// *****************************************
        /// *ColumnName               | ColumnValue * 
        /// **************************|**************
        /// *[dbo].[Education].[School]|屏東科技大學*
        /// *****************************************

        /// <summary>
        ///透過Table and 欄位查詢對應的ID
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="da"></param>
        /// <returns></returns>
        private List<string> SelectIdbyResult(DataTable dt, SqlDataAdapter da)
        {
            List<string> idList = new List<string>();
            da.SelectCommand.Parameters.Clear();
            da.SelectCommand.CommandText = string.Empty;
            da.SelectCommand.CommandType = CommandType.Text;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                string[] tableName = dr[0].ToString().Split('.');
                if (tableName[1] == "[Contact_Status_List]" || tableName[1] == "[Member]")
                {
                    continue;
                }

                if (tableName[1] == "[Code]" || tableName[1] == "[Contact_Situation]" || tableName[1] == "[Contact_Info]" || tableName[1] == "[Interview_Info]")
                {
                    if (i == dt.Rows.Count - 1)
                    {
                        da.SelectCommand.CommandText += @"select Contact_Id from " + tableName[1] + " where " + tableName[2] + " = @value" + i;
                    }
                    else
                    {
                        da.SelectCommand.CommandText += @"select Contact_Id from " + tableName[1] + " where " + tableName[2] + " = @value" + i + " UNION ";
                    }

                    da.SelectCommand.Parameters.Add("@value" + i, SqlDbType.NVarChar).Value = dr[1].ToString();
                }
                else
                {
                    if (i == dt.Rows.Count - 1)
                    {
                        da.SelectCommand.CommandText += @"select b.Contact_Id from " + tableName[1] + " a,Interview_Info b where a.Interview_Id = b.Interview_Id and a." + tableName[2] + " = @value" + i;
                    }
                    else
                    {
                        da.SelectCommand.CommandText += @"select b.Contact_Id from " + tableName[1] + " a,Interview_Info b where a.Interview_Id = b.Interview_Id and a." + tableName[2] + " = @value" + i + " UNION ";
                    }

                    da.SelectCommand.Parameters.Add("@value" + i, SqlDbType.NVarChar).Value = dr[1].ToString();
                }
            }

            Common.GetInstance().ClearDataTable(dt);
            if (da.SelectCommand.CommandText.EndsWith(" UNION "))
            {
                da.SelectCommand.CommandText = da.SelectCommand.CommandText.Remove(da.SelectCommand.CommandText.Length - 7);
            }

            da.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    idList.Add(dr[0].ToString());
                }
            }

            return idList;
        }

        /// <summary>
        /// 根據ID撈出人才資料
        /// </summary>
        /// <param name="idList"></param>
        /// <returns></returns>
        public DataTable SelectTalentInfoById(List<string> idList)
        {
            ErrorMessage = string.Empty;
            string select = string.Empty;
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            dt.Columns.Add("Contact_Id");
            dt.Columns.Add("Name");
            dt.Columns.Add("Code_Id");
            dt.Columns.Add("Contact_Date");
            dt.Columns.Add("Contact_Status");
            dt.Columns.Add("Remarks");
            dt.Columns.Add("Interview_Date");
            dt.Columns.Add("UpdateTime");
            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    foreach (string id in idList)
                    {
                        ds.Tables.Clear();
                        da.SelectCommand.Parameters.Clear();
                        da.SelectCommand.CommandText = @"select top 1 Contact_Id,Name,CONVERT(varchar(100), UpdateTime, 111) UpdateTime from Contact_Info where Contact_Id = @Contact_Id ";
                        da.SelectCommand.CommandText += @"select Code_Id from Code where Contact_Id=@Contact_Id ";
                        da.SelectCommand.CommandText += @"select top 2 CONVERT(varchar(100), Contact_Date, 111) Contact_Date,Contact_Status,Remarks from Contact_Situation where Contact_Id = @Contact_Id order by Contact_Date desc ";
                        da.SelectCommand.CommandText += @"select top 2 CONVERT(varchar(100), Interview_Date, 111) Interview_Date from Interview_Info where Contact_Id = @Contact_Id order by Interview_Date desc";
                        da.SelectCommand.Parameters.Add("@Contact_Id", SqlDbType.Int).Value = id;
                        da.Fill(ds);
                        CombinationGrid(ds, dt);
                    }
                }
                return dt;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return new DataTable();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 根據聯繫ID查詢相關的聯繫基本資料與聯繫狀況
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public DataSet SelectContactSituationDataById(string id)
        {
            ErrorMessage = string.Empty;
            DataSet ds = new DataSet();
            string select = @"select [Name],[Sex],[Mail],[CellPhone],[Cooperation_Mode],[Status],[Place],[Skill],[Year] from Contact_Info a where a.Contact_Id = @id
                              select CONVERT(varchar(100), a.Contact_Date, 111) Contact_Date,a.Contact_Status,a.Remarks from Contact_Situation a where a.Contact_Id = @id order by Contact_Date desc
                              select a.Code_Id from Code a where a.Contact_Id = @id
                              select CONVERT(varchar(100), UpdateTime, 20) UpdateTime from Contact_Info where Contact_Id = @id
                              select Interview_Id from Interview_Info where Contact_Id = @id";
            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    da.SelectCommand.Parameters.Add("@id", SqlDbType.Int).Value = id;
                    da.Fill(ds);
                }

                return ds;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return new DataSet();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        public DataTable SelectContactSituationInfoByCode(List<string> CodeList)
        {
            DataTable dt = new DataTable();
            string whereCode = string.Empty;
            try
            {
                for (int i = 0; i < CodeList.Count(); i++)
                {
                    whereCode += "@Code" + i + ", ";
                }

                whereCode = whereCode.Substring(0, whereCode.Length - 2);
                string select = @"select a.Name,a.Code,b.Contact_Date,b.Contact_Status,b.Remarks,a.UpdateTime 
                              from Talent_Info a left join Contact_Situation b on a.Code = b.Code 
                              and a.Code in (select code from Talent_Info where code in (" + whereCode + ") and Visible = 'true')";
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    for (int i = 0; i < CodeList.Count(); i++)
                    {
                        da.SelectCommand.Parameters.Add("@Code" + i, SqlDbType.Int).Value = CodeList[i];
                    }

                    da.Fill(dt);
                    return dt;
                }
            }
            catch
            {
                return new DataTable();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 根據面談ID與附加檔案模式來查詢符合的資料
        /// </summary>
        /// <param name="interviewId">面談ID</param>
        /// <param name="filesMode">附加檔案模式 EX：面談基本資料，專案經驗</param>
        /// <returns></returns>
        public DataTable SelectFiles(string interviewId, string filesMode)
        {
            ErrorMessage = string.Empty;
            DataTable dt = new DataTable();
            string select = @"select File_Path,Belong from Files where Interview_Id = @Interview_Id and Belong = @Belong";
            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    da.SelectCommand.Parameters.Add("@Interview_Id", SqlDbType.Int).Value = interviewId;
                    da.SelectCommand.Parameters.Add("@Belong", SqlDbType.NVarChar).Value = filesMode;
                    da.Fill(dt);
                }

                return dt;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return new DataTable();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        ///根據面談資料ID查詢面談資料
        /// </summary>
        /// <param name="interviewId"></param>
        /// <returns></returns>
        public DataSet SelectInterviewDataById(string interviewId)
        {
            ErrorMessage = string.Empty;
            DataSet ds = new DataSet();
            string SQLStr = @"select [Vacancies],CONVERT(varchar(100), Interview_Date, 111)Interview_Date,[Name],[Sex],CONVERT(varchar(100), [Birthday], 111) Birthday,[Married],[Mail],[Adress],[CellPhone],[Image]
                                    ,[Expertise_Language],[Expertise_Tools],[Expertise_Tools_Framwork],[Expertise_Devops],[Expertise_OS],[Expertise_BigData],[Expertise_DataBase]
                                    ,[Expertise_Certification],[IsStudy],[IsService],[Relatives_Relationship],[Relatives_Name],[Care_Work],[Hope_Salary]
                                    ,[When_Report],[Advantage],[Disadvantages],[Hobby],[Attract_Reason],[Future_Goal],[Hope_Supervisor]
                                    ,[Hope_Promise],[Introduction],[During_Service],[Exemption_Reason],[Urgent_Contact_Person],[Urgent_Relationship]
                                    ,[Urgent_CellPhone],[Education],[Language],[Work_Experience] from Interview_Info where Interview_Id=@interviewId
                              select Interviewer,Result from Interview_Comments  where Interview_Id=@interviewId
                              select Appointment,Results_Remark from Interview_Info where Interview_Id=@interviewId
                              select [Company],[Project_Name],[OS],[Database],[Position],[Language],[Tools],[Description],[Start_End_Date] from Project_Experience  where Interview_Id=@interviewId";
            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(SQLStr, ScConnection))
                {
                    da.SelectCommand.Parameters.Add("@interviewId", SqlDbType.Int).Value = interviewId;
                    da.Fill(ds);
                }
                return ds;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return new DataSet();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 搜尋使用者帳號資訊，hr帳號要排除
        /// </summary>
        /// <returns>使用者帳號資訊</returns>
        public DataTable SelectMemberInfo()
        {
            ErrorMessage = string.Empty;
            DataTable dt = new DataTable();
            string select = @"select Account,States from Member where account != 'hr@is-land.com.tw'";
            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    da.Fill(dt);
                    if (dt.Rows.Count == 0)
                    {
                        return new DataTable();
                    }
                }

                return dt;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return new DataTable();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 根據帳號搜尋使用者資訊
        /// </summary>
        /// <param name="account">帳號</param>
        /// <returns>帳號資訊EX：hr@is-land.com.tw,啟用</returns>
        public DataTable SelectMemberInfoByAccount(string account)
        {
            ErrorMessage = string.Empty;
            DataTable dt = new DataTable();
            string select = @"select Account,States from Member where Account=@account";
            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    da.SelectCommand.Parameters.Add("@account", SqlDbType.NVarChar).Value = account;
                    da.Fill(dt);
                    if (dt.Rows.Count == 0)
                    {
                        return new DataTable();
                    }
                }

                return dt;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return new DataTable();
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }
    }
}
