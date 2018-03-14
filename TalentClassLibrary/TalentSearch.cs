using ServiceStack;
using ShareClassLibrary;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TalentClassLibrary.Model;

namespace TalentClassLibrary
{
    public class TalentSearch : SQL
    {
        private static TalentSearch talentSearch = new TalentSearch();

        public static TalentSearch GetInstance() => talentSearch;

        /// <summary>
        /// 紀錄目前發生
        /// </summary>
        public string ErrorMessage { get; set; }

       /// <summary>
       /// 儲存面談資料
       /// </summary>
       /// <param name="saveData">面談資料</param>
       /// <param name="interviewId">面談ID</param>
       /// <param name="dbPath">資料庫的圖片路徑</param>
       /// <param name="uiPath">UI的圖片路徑</param>
       /// <returns></returns>
        public string SaveInterviewData(DataSet saveData, string interviewId, string dbPath, string uiPath)
        {
            if (!Talent.GetInstance().ValidInterviewIdIsAppear(interviewId))
            {
                return "沒有對應的面試基本資料";
            }

            string path = TalentFiles.GetInstance().UpdateImage(dbPath, uiPath, interviewId);
            if (path.Equals("上傳失敗") || path.Equals("不存在的路徑") || path.Equals("檔案刪除失敗"))
            {
                throw new Exception(path);
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
                    ////儲存面談基本資料
                    foreach (DataRow dr in saveData.Tables[1].Rows)
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
                    ////儲存專案經驗資料
                    cmd.CommandText = @"delete from [Project_Experience] where Interview_Id=@interviewId";
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = @"insert into [Project_Experience] ([Company],[Project_Name],[OS],[Database],[Position],[Language],[Tools],[Description],[Start_End_Date],[Interview_Id])
                                        values (@Company,@Project_Name,@OS,@Database,@Position,@Language,@Tools,@Description,@Start_End_Date,@Interview_Id)";
                    foreach (DataRow dr in saveData.Tables[2].Rows)
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
                    ////儲存面談結果資料
                    cmd.CommandText = @"update Interview_Info set Appointment=@Appointment,Results_Remark=@Results_Remark where Interview_Id=@Interview_Id";
                    foreach (DataRow dr in saveData.Tables[3].Rows)
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
        /// 將Excel的資料寫入資料庫
        /// </summary>
        /// <param name="contactSituationList"></param>
        /// <returns></returns>
        public string InsertContactSituationInfoData(List<ContactSituation> contactSituationList)
        {
            string contact_Id = string.Empty;
            string insert = string.Empty;
            try
            {
                using (SqlCommand cmd = new SqlCommand(insert, ScConnection, StTransaction))
                {
                    foreach (ContactSituation contactSituation in contactSituationList)
                    {
                        cmd.CommandText = @"insert into Contact_Info ([Name],[Sex],[Mail],[CellPhone],[UpdateTime],[Cooperation_Mode],[Status],[Place],[Skill],[Year])
                              output inserted.Contact_Id
                              values(@name,@sex,@mail,@cellPhone,GETDATE(),@cooperationMode,@status,@place,@skill,@year)";
                        cmd.Parameters.Clear();
                        ////新增聯繫資料
                        cmd.Parameters.Add("@name", SqlDbType.NVarChar).Value = contactSituation.Info.Name ?? string.Empty;
                        cmd.Parameters.Add("@sex", SqlDbType.NVarChar).Value = contactSituation.Info.Sex ?? string.Empty;
                        cmd.Parameters.Add("@mail", SqlDbType.VarChar).Value = contactSituation.Info.Mail ?? string.Empty;
                        cmd.Parameters.Add("@cellPhone", SqlDbType.VarChar).Value = contactSituation.Info.CellPhone ?? string.Empty;
                        cmd.Parameters.Add("@cooperationMode", SqlDbType.NVarChar).Value = "皆可";
                        cmd.Parameters.Add("@status", SqlDbType.NVarChar).Value = string.Empty;
                        cmd.Parameters.Add("@place", SqlDbType.NVarChar).Value = contactSituation.Info.Place ?? string.Empty;
                        cmd.Parameters.Add("@skill", SqlDbType.NVarChar).Value = contactSituation.Info.Skill ?? string.Empty;
                        cmd.Parameters.Add("@year", SqlDbType.VarChar).Value = string.Empty;
                        contact_Id = cmd.ExecuteScalar().ToString();
                        ////新增聯繫狀況
                        cmd.CommandText = @"insert into Contact_Situation (Contact_Status,Remarks,Contact_Date,Contact_Id)
                                        values (@Contact_Status,@Remarks,@Contact_Date,@Contact_Id)";
                        foreach (ContactStatus contactstatus in contactSituation.Status)
                        {
                            cmd.Parameters.Clear();
                            cmd.Parameters.Add(@"Contact_Status", SqlDbType.NVarChar).Value = contactstatus.Contact_Status ?? "人資系統資料";
                            cmd.Parameters.Add(@"Remarks", SqlDbType.NVarChar).Value = contactstatus.Remarks ?? string.Empty;
                            cmd.Parameters.Add(@"Contact_Date", SqlDbType.DateTime).Value = string.IsNullOrEmpty(contactstatus.Contact_Date) ? "2000/1/1" : contactstatus.Contact_Date;
                            cmd.Parameters.Add(@"Contact_Id", SqlDbType.Int).Value = contact_Id;
                            cmd.ExecuteNonQuery();
                        }
                        ////新增代碼
                        cmd.CommandText = @"insert into Code (Code_Id,Contact_Id)
                                        values (@Code_Id,@Contact_Id)";
                        if (string.IsNullOrEmpty(contactSituation.Code))
                        {
                            continue;
                        }

                        string[] codeList = contactSituation.Code.Trim().Split('\n');

                        foreach (string code in codeList)
                        {
                            cmd.Parameters.Clear();
                            cmd.Parameters.Add(@"Code_Id", SqlDbType.VarChar).Value = code;
                            cmd.Parameters.Add(@"Contact_Id", SqlDbType.Int).Value = contact_Id;
                            cmd.ExecuteNonQuery();
                        }
                    }

                    this.CommitTransaction();
                    return "匯入成功";
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                this.RollbackTransaction();
                return "匯入失敗";
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        /// <summary>
        /// 進入查詢頁面時，顯示最新(依造最後編輯日降冪排序)的15筆資料
        /// </summary>
        /// <returns></returns>
        public DataTable SelectTop15()
        {
            ErrorMessage = string.Empty;
            DataTable dt = new DataTable();
            string select = @"select Contact_Id,Name,Code_Id,CONVERT(varchar(100), Contact_Date, 111) Contact_Date,Contact_Status,Remarks,CONVERT(varchar(100), Interview_Date, 111) Interview_Date,CONVERT(varchar(100), UpdateTime, 120) UpdateTime 
                              from FilterTable where Contact_Id in(
                              select a.Contact_Id
                              from (select distinct top 15 Contact_Id,UpdateTime from FilterTable order by UpdateTime desc) a)";
            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
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
        /// 根據條件查詢符合的資料
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
        public DataTable SelectIdByFilter(string keyWords, string places, string expertises, string cooperationMode, string states, string startEditDate, string endEditDate, string isInterview, string interviewResult, string startInterviewDate, string endInterviewDate)
        {
            ErrorMessage = string.Empty;
            DataTable dt = new DataTable();
            string select = @"select Contact_Id,Name,Code_Id,CONVERT(varchar(100), Contact_Date, 111) Contact_Date,Contact_Status,Remarks,CONVERT(varchar(100), Interview_Date, 111) Interview_Date,CONVERT(varchar(100), UpdateTime, 120) UpdateTime from FilterTable where 1=1 ";
            try
            {
                if (Valid.GetInstance().ValidDateRange(startEditDate, endEditDate) != string.Empty)
                {
                    ErrorMessage = "最後編輯日之日期格式或者是日期區間不正確";
                    return new DataTable();
                }

                if (Valid.GetInstance().ValidDateRange(startInterviewDate, endInterviewDate) != string.Empty)
                {
                    ErrorMessage = "面談日期之日期格式或者是日期區間不正確";
                    return new DataTable();
                }

                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    this.CombinationWhereByContactFilter(places, expertises, cooperationMode, states, startEditDate, endEditDate, da);
                    this.CombinationWhereByInterviewFilter(isInterview, interviewResult, startInterviewDate, endInterviewDate, da);
                    this.CombinationWhereByKeyWord(keyWords, da);
                    da.SelectCommand.CommandText += " order by UpdateTime desc";
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
        /// 將查詢結果組合成符合格式的DataTable
        /// </summary>
        /// <param name="dt">查詢決果</param>
        /// <returns></returns>
        public DataTable CombinationGrid(DataTable dt)
        {
            ErrorMessage = string.Empty;
            try
            {
                DataTable dataTable = dt.Clone();
                var idList = dt.AsEnumerable().Select(x => x.Field<int>("Contact_Id")).Distinct().ToList();
                List<SearchResult> searchResultList = new List<SearchResult>();
                for (int i = 0; i < idList.Count; i++)
                {
                    SearchResult searchResult = new SearchResult
                    {
                        Id = idList[i].ToString(),
                    };

                    DataSet dataSet = new DataSet();

                    DataTable contactDt = (from row in dt.AsEnumerable()
                                           where row.Field<int>("Contact_Id") == idList[i]
                                           select new
                                           {
                                               Contact_Id = row.Field<int>("Contact_Id"),
                                               Name = row.Field<string>("Name"),
                                               UpdateTime = row.Field<string>("UpdateTime")
                                           }).Distinct().LinqQueryToDataTable();
                    DataTable codeDt = (from row in dt.AsEnumerable()
                                        where row.Field<int>("Contact_Id") == idList[i]
                                        select new
                                        {
                                            Code_Id = row.Field<string>("Code_Id")
                                        }).Distinct().LinqQueryToDataTable();
                    DataTable statusDt = (from row in dt.AsEnumerable()
                                          where row.Field<int>("Contact_Id") == idList[i]
                                          select new
                                          {
                                              Contact_Date = row.Field<string>("Contact_Date"),
                                              Contact_Status = row.Field<string>("Contact_Status"),
                                              Remarks = row.Field<string>("Remarks")
                                          }).Distinct().OrderByDescending(x => x.Contact_Date).Take(2).LinqQueryToDataTable();
                    DataTable interviewDateDt = (from row in dt.AsEnumerable()
                                                 where row.Field<int>("Contact_Id") == idList[i]
                                                 select new
                                                 {
                                                     Interview_Date = row.Field<string>("Interview_Date")
                                                 }).Distinct().OrderByDescending(x => x.Interview_Date).Take(2).LinqQueryToDataTable();
                    dataSet.Tables.Add(contactDt);
                    dataSet.Tables.Add(codeDt);
                    dataSet.Tables.Add(statusDt);
                    dataSet.Tables.Add(interviewDateDt);
                    Talent.CombinationGrid(dataSet, dataTable);
                }
                return dataTable;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "查詢時，發生錯誤";
                return new DataTable();
            }
        }

        /// <summary>
        /// 組合面談資訊條件的where語法
        /// </summary>
        /// <param name="isInterview">是否已面談，值為"已面談","未面談","不限"</param>
        /// <param name="interviewResult">面談結果，值為"錄用","不錄用","暫保留","不限"</param>
        /// <param name="startInterviewDate">起始日期，日期格式</param>
        /// <param name="endInterviewDate">結束日期，日期格式</param>
        /// <param name="da"></param>
        /// <returns></returns>
        private SqlDataAdapter CombinationWhereByInterviewFilter(string isInterview, string interviewResult, string startInterviewDate, string endInterviewDate, SqlDataAdapter da)
        {
            ////是否已面談
            if (isInterview.Equals("已面談"))
            {
                da.SelectCommand.CommandText += @" and Appointment is not null and Appointment !=''";
            }
            else if (isInterview.Equals("未面談"))
            {
                da.SelectCommand.CommandText += @" and Appointment is null or Appointment =''";
            }

            ////面談結果
            if (!interviewResult.Equals("不限"))
            {
                da.SelectCommand.CommandText += @" and ISNULL(Appointment,'NA') = ISNULL(ISNULL(@interviewResult,Appointment),'NA')";
                da.SelectCommand.Parameters.Add("@interviewResult", SqlDbType.NVarChar).Value = Common.GetInstance().ValueIsNullOrEmpty(interviewResult);
            }

            ////面談日期
            if (!string.IsNullOrEmpty(startInterviewDate))
            {
                da.SelectCommand.CommandText += @" AND Interview_Date >= ISNULL(@startInterviewDate, Interview_Date)";
                da.SelectCommand.Parameters.Add("@startInterviewDate", SqlDbType.DateTime).Value = Common.GetInstance().ValueIsNullOrEmpty(startInterviewDate);
            }

            if (!string.IsNullOrEmpty(endInterviewDate))
            {
                da.SelectCommand.CommandText += @" and Interview_Date <= ISNULL(@endInterviewDate, Interview_Date) ";
                da.SelectCommand.Parameters.Add("@endInterviewDate", SqlDbType.DateTime).Value = Common.GetInstance().ValueIsNullOrEmpty(endInterviewDate);
            }
            return da;
        }

        /// <summary>
        /// 組合聯繫資訊的where語法
        /// </summary>
        /// <param name="places">地點，多筆請用","隔開</param>
        /// <param name="expertises">技能，多筆請用","隔開</param>
        /// <param name="cooperationMode">合作模式</param>
        /// <param name="states">聯繫狀態</param>
        /// <param name="startEditDate">起始日期，日期格式</param>
        /// <param name="endEditDate">結束日期，日期格式</param>
        /// <param name="da"></param>
        /// <returns></returns>
        private SqlDataAdapter CombinationWhereByContactFilter(string places, string expertises, string cooperationMode, string states, string startEditDate, string endEditDate, SqlDataAdapter da)
        {
            CombinationWhereConditionStates(states, da);
            CombinationWhereConditionStartEditDate(startEditDate, da);
            CombinationWhereConditionEndEditDate(endEditDate, da);
            CombinationWhereConditionCooperationMode(cooperationMode, da);
            CombinationWhereConditionPlaces(places, da);
            CombinationWhereConditionExpertises(expertises, da);

            return da;
        }

        #region CombinationWhereByContactFilter function
        private SqlDataAdapter CombinationWhereConditionStates(string states, SqlDataAdapter da)
        {
            if (!DBNull.Value.Equals(this.ValueIsAny(states)))
            {
                da.SelectCommand.CommandText += @" and ISNULL(Status,'NA') = ISNULL(ISNULL(@status,Status),'NA') ";

                da.SelectCommand.Parameters.Add("@status", SqlDbType.NVarChar).Value = this.ValueIsAny(states);
            }
            return da;
        }

        private SqlDataAdapter CombinationWhereConditionStartEditDate( string startEditDate, SqlDataAdapter da)
        {
            if (!DBNull.Value.Equals(Common.GetInstance().ValueIsNullOrEmpty(startEditDate)))
            {
                da.SelectCommand.CommandText += @" and UpdateTime >= ISNULL(@startEditDate, UpdateTime) ";

                da.SelectCommand.Parameters.Add("@startEditDate", SqlDbType.DateTime).Value = Common.GetInstance().ValueIsNullOrEmpty(startEditDate);
            }
            return da;
        }

        private SqlDataAdapter CombinationWhereConditionEndEditDate(string endEditDate, SqlDataAdapter da)
        {
            if (!DBNull.Value.Equals(Common.GetInstance().ValueIsNullOrEmpty(endEditDate)))
            {
                da.SelectCommand.CommandText += @" and UpdateTime <= ISNULL(@endEditDate, UpdateTime) ";

                da.SelectCommand.Parameters.Add("@endEditDate", SqlDbType.DateTime).Value = Common.GetInstance().ValueIsNullOrEmpty(endEditDate);
            }
            return da;
        }

        private SqlDataAdapter CombinationWhereConditionCooperationMode(string cooperationMode, SqlDataAdapter da)
        {
            if (!DBNull.Value.Equals(this.ValueIsAny(cooperationMode)))
            {
                da.SelectCommand.CommandText += @" and ISNULL(Cooperation_Mode,'NA') = ISNULL(ISNULL(@CooperationMode,Cooperation_Mode),'NA') ";
                da.SelectCommand.Parameters.Add("@CooperationMode", SqlDbType.NChar).Value = this.ValueIsAny(cooperationMode);

                ////如果合作模式為"全職"or"合約"，則值為"皆可"也要被查詢出來
                if (cooperationMode.Equals("全職") || cooperationMode.Equals("合約"))
                {
                    da.SelectCommand.CommandText += @" or ISNULL(Cooperation_Mode,'NA') = ISNULL(ISNULL(@CooperationMode1,Cooperation_Mode),'NA')";
                    da.SelectCommand.Parameters.Add("@CooperationMode1", SqlDbType.NChar).Value = Common.GetInstance().ValueIsNullOrEmpty("皆可");
                }
            }
            return da;
        }

        private SqlDataAdapter CombinationWhereConditionPlaces(string places, SqlDataAdapter da)
        {
            if (!string.IsNullOrEmpty(places.Trim()))
            {
                string[] place = places.Split(',');

                ////多筆地點
                for (int i = 0; i < place.Length; i++)
                {
                    da.SelectCommand.CommandText += (i == 0) ? @" and " : @" or ";
                    da.SelectCommand.CommandText += @"ISNULL(Place,'NA') LIKE ISNULL(ISNULL(@place" + (i + 1) + ", Place),'NA')";
                    da.SelectCommand.Parameters.Add("@place" + (i + 1), SqlDbType.NVarChar).Value = Common.GetInstance().ValueIsNullOrEmpty("%" + place[i] + "%");
                }
            }
            return da;
        }

        private SqlDataAdapter CombinationWhereConditionExpertises(string expertises, SqlDataAdapter da)
        {

            if (!string.IsNullOrEmpty(expertises.Trim()))
            {
                string[] expertise = expertises.Split(',');

                ////多筆技能
                for (int i = 0; i < expertise.Length; i++)
                {

                    da.SelectCommand.CommandText += (i == 0) ? @" and " : @" or ";
                    da.SelectCommand.CommandText += @"ISNULL(Skill,'NA') Like ISNULL(ISNULL(@skill" + (i + 1) + ", Skill),'NA')";
                    da.SelectCommand.Parameters.Add("@skill" + (i + 1), SqlDbType.NVarChar).Value = Common.GetInstance().ValueIsNullOrEmpty("%" + expertise[i] + "%");

                }
            }

            return da;
        }
        #endregion

        /// <summary>
        /// 組合關鍵字的where語法
        /// </summary>
        /// <param name="keyWords">關鍵字，如有多個關鍵字請用","隔開</param>
        /// <param name="da"></param>
        /// <returns></returns>
        private SqlDataAdapter CombinationWhereByKeyWord(string keyWords, SqlDataAdapter da)
        {
            if (!string.IsNullOrEmpty(keyWords))
            {
                string[] keyWord = keyWords.Split(',');
                for (int i = 0; i < keyWord.Length; i++)
                {
                    if (string.IsNullOrEmpty(keyWord[i]))
                    {
                        continue;
                    }

                    da.SelectCommand.CommandText += (i == 0) ? @" and (" : @" or ";
                    da.SelectCommand.CommandText += @" Name like @word" + (i + 1) + " or Code_Id like @word" + (i + 1) + " or Remarks like @word" + (i + 1);
                    da.SelectCommand.Parameters.Add("@word" + (i + 1), SqlDbType.NVarChar).Value = "%" + keyWord[i] + "%";
                }

                da.SelectCommand.CommandText += ")";
            }
            return da;
        }

        /// <summary>
        /// 如果值為"不限"，NULL，空，則回傳DBNULL
        /// <param name="value">要判斷的值</param>
        /// <returns>會傳值或者是DBNULL</returns>
        private object ValueIsAny(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return DBNull.Value;
            }
            else
            {
                if (value.Equals("不限"))
                {
                    return DBNull.Value;
                }

                return value;
            }
        }
    }
}
