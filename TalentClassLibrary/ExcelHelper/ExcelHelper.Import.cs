using Newtonsoft.Json;
using ServiceStack;
using ShareClassLibrary;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using TalentClassLibrary.Model;

namespace TalentClassLibrary
{
    /// <summary>
    /// 處理Excel的類別
    /// </summary>
    public partial class ExcelHelper
    {
        /// <summary>
        /// 匯入面談資料
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public DataSet ImportInterviewData(string path)
        {
            ErrorMessage = string.Empty;
            DataSet ds = new DataSet();
            try
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(path);
                ////面談基本資料
                Worksheet sheet = workbook.Worksheets["人事資料"];               
                ////面談結果
                ds.Tables.Add(this.ReadInterviewInfoSheet(sheet));               
                ds.Tables.Add(new List<InterviewComments>().ListToDataTable());
                ds.Tables.Add(new List<InterviewResult>().ListToDataTable());              
                ////專案經驗
                Worksheet sheet1 = workbook.Worksheets["專案經驗"];
                ds.Tables.Add(this.ReadProjectExperienceSheet(sheet1));

                return ds;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "匯入失敗";
                ds.Clear();
                return ds;
            }
        }

        /// <summary>
        /// 匯入舊版資料
        /// </summary>
        /// <param name="path"></param>
        public List<ContactSituation> ImportOldTalent(string path)
        {
            ErrorMessage = string.Empty;
            List<ContactSituation> contactSituationList = new List<ContactSituation>();
            try
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(path);
                Worksheet sheet = workbook.Worksheets[0];
                DataTable dt = sheet.ExportDataTable();
                //if (!TalentValid.GetInstance().ValidIsOldTalentFormat(dt.Columns))
                if (!TalentValid.GetInstance().ValidTalentFormat(dt.Columns,true))
                {
                    ErrorMessage = "excel格式不符";
                    return contactSituationList;
                }

                if (dt.Rows.Count == 0)
                {
                    ErrorMessage = "空的excel";
                    return contactSituationList;
                }

                foreach (DataRow dr in dt.Rows)
                {
                    ContactSituation contactSituation = new ContactSituation();
                    ContactInfo contactInfo = new ContactInfo();
                    ContactStatus contactStatus = new ContactStatus();
                    for (int i = 0; i < 10; i++)
                    {
                        switch (dt.Columns[i].ToString().Trim())
                        {
                            case "姓名":
                                contactInfo.Name = dr[i].ToString().Trim();
                                break;
                            case "性別":
                                contactInfo.Sex = dr[i].ToString().Trim();
                                break;
                            case "手機":
                                contactInfo.CellPhone = dr[i].ToString().Trim();
                                break;
                            case "郵件":
                                contactInfo.Mail = dr[i].ToString().Trim();
                                break;
                            case "學校":
                                if (!string.IsNullOrEmpty(dr[i].ToString().Trim()))
                                {
                                    contactStatus.Remarks += "學校\n" + dr[i].ToString().Trim() + "\n";
                                }

                                break;
                            case "地區":
                                contactInfo.Place = dr[i].ToString().Trim();
                                break;
                            case "最後編輯時間":
                                contactStatus.Contact_Date = dr[i].ToString().Trim();
                                break;
                            case "專長":
                                contactInfo.Skill = dr[i].ToString().Trim();
                                break;
                            case "互動":
                                if (!string.IsNullOrEmpty(dr[i].ToString().Trim()))
                                {
                                    contactStatus.Remarks += "互動\n" + dr[i].ToString().Trim() + "\n";
                                }

                                break;
                            case "評價":
                                if (!string.IsNullOrEmpty(dr[i].ToString().Trim()))
                                {
                                    contactStatus.Remarks += "評價\n" + dr[i].ToString().Trim() + "\n";
                                }

                                break;
                        }
                    }
                    contactSituation.Info = contactInfo;
                    contactSituation.Status = new List<ContactStatus> { contactStatus };
                    contactSituationList.Add(contactSituation);
                }

                return contactSituationList;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                contactSituationList.Clear();
                ErrorMessage = "讀取Excel發生錯誤";
                return contactSituationList;
            }
        }

        /// <summary>
        /// 匯入新版資料
        /// </summary>
        /// <param name="path"></param>
        public List<ContactSituation> ImportNewTalent(string path)
        {
            ErrorMessage = string.Empty;
            string msg = string.Empty;
            List<string> checkCodeIsRepeat = new List<string>(); ////檢查Excel內部的代碼是否有重複
            List<ContactSituation> contactSituationList = new List<ContactSituation>();
            try
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(path);
                Worksheet sheet = workbook.Worksheets[0];
                DataTable dt = sheet.ExportDataTable();

                //if(!TalentValid.GetInstance().ValidIsNewTalentFormat(dt.Columns))
                if (!TalentValid.GetInstance().ValidTalentFormat(dt.Columns,false))
                {
                    ErrorMessage = "excel格式不符";
                    return contactSituationList;
                }

                if (dt.Rows.Count == 0)
                {
                    ErrorMessage = "空的excel";
                    return contactSituationList;
                }

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ContactSituation contactSituation = new ContactSituation();
                    ContactInfo contactInfo = new ContactInfo();
                    List<ContactStatus> contactStatusList = new List<ContactStatus>();
                    for (int j = 0; j < 29; j++)
                    {
                        ////先不處理聯繫狀況
                        if (j == 3 || j == 4 || j == 5)
                        {
                            continue;
                        }

                        switch (dt.Columns[j].ToString().Trim())
                        {
                            case "姓名":
                                contactInfo.Name = dt.Rows[i].ItemArray[j].ToString().Trim();
                                break;
                            case "地點":
                                contactInfo.Place = dt.Rows[i].ItemArray[j].ToString().Trim();
                                break;
                            case "1111/104代碼":
                                ////檢查Excel內部的代碼是否有重複
                                string[] codeList = dt.Rows[i].ItemArray[j].ToString().Trim().Split('\n');
                                foreach (string code in codeList)
                                {
                                    if (checkCodeIsRepeat.Contains(code))
                                    {
                                        contactSituationList.Clear();
                                        ErrorMessage = "第" + (i + 1) + "行" + code + "重複\n請檢查Excel";
                                        return contactSituationList;
                                    }

                                    checkCodeIsRepeat.Add(code);
                                }

                                contactSituation.Code = dt.Rows[i].ItemArray[j].ToString().Trim();
                                break;
                            case "JAVA":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "JSP":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "Android APP":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "ASP":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "C/C++":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "C#":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "ASP.NET":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "VB.NET":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "VB6":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "HTML":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "Javascript":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "Bootstrap":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "Delphi":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "PHP":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "研替":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "Hadoop":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "ETL":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "R":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "notes":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "UI/UX":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "資料庫":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                            case "Linux":
                                if (!string.IsNullOrEmpty(dt.Rows[i].ItemArray[j].ToString().Trim()))
                                    contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString().Trim() + ",";
                                break;
                        }
                    }

                    ////聯繫狀況
                    do
                    {
                        ContactStatus contactStatus = new ContactStatus();
                        for (int z = 3; z <= 5; z++)
                        {
                            switch (dt.Columns[z].ToString().Trim())
                            {
                                case "日期":
                                    msg = Valid.GetInstance().ValidDateFormat(dt.Rows[i].ItemArray[z].ToString().Trim());
                                    if (msg != string.Empty)
                                    {
                                        contactSituationList.Clear();
                                        ErrorMessage = "第" + (i + 1) + "行" + msg + "\n請檢查Excel";
                                        return contactSituationList;
                                    }

                                    contactStatus.Contact_Date = dt.Rows[i].ItemArray[z].ToString().Trim();
                                    break;
                                case "聯絡狀況":
                                    if (string.IsNullOrEmpty(dt.Rows[i].ItemArray[z].ToString().Trim()))
                                    {
                                        contactStatus.Contact_Status = "(無)";
                                    }
                                    else
                                    {
                                        msg = TalentValid.GetInstance().ValidContactStatus(dt.Rows[i].ItemArray[z].ToString().Trim());
                                        if (msg != string.Empty)
                                        {
                                            contactSituationList.Clear();
                                            ErrorMessage = "第" + (i + 1) + "行" + msg + "\n請檢查Excel";
                                            return contactSituationList;
                                        }

                                        contactStatus.Contact_Status = dt.Rows[i].ItemArray[z].ToString().Trim();
                                    }
                                    break;
                                case "說明":
                                    contactStatus.Remarks = dt.Rows[i].ItemArray[z].ToString().Trim();
                                    break;
                            }
                        }

                        contactStatusList.Add(contactStatus);

                        i++;
                        if (i >= dt.Rows.Count)
                        {
                            break;
                        }
                    } while (string.IsNullOrEmpty(dt.Rows[i].ItemArray[0].ToString().Trim()) && string.IsNullOrEmpty(dt.Rows[i].ItemArray[1].ToString().Trim()));

                    i--; ////因為在do迴圈有先i++判斷下一列的姓名與代碼是否有值，因此在跳出迴圈時，要把它減回來
                    contactInfo.Skill = contactInfo.Skill.RemoveEndWithDelimiter(",");
                    contactSituation.Info = contactInfo;
                    contactSituation.Status = contactStatusList;
                    contactSituationList.Add(contactSituation);
                }

                msg = Talent.GetInstance().ValidCodeIsRepeat(checkCodeIsRepeat);
                if (msg != string.Empty)
                {
                    contactSituationList.Clear();
                    ErrorMessage = msg + "\n請檢查Excel";
                    return contactSituationList;
                }

                return contactSituationList;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                contactSituationList.Clear();
                ErrorMessage = "讀取Excel發生錯誤";
                return contactSituationList;
            }
        }
    }
}
