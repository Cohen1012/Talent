using ShareClassLibrary;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TalentClassLibrary.Model;

namespace TalentClassLibrary
{
    /// <summary>
    /// 一般功能類別
    /// </summary>
    public class TalentCommon
    {
        private static TalentCommon talentCommon = new TalentCommon();

        public static TalentCommon GetInstance() => talentCommon;

        /// <summary>
        /// 根據聯繫ID匯出聯繫資料
        /// </summary>
        /// <param name="contactIdList"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        public string ExportContactSituationDataByContactId(List<string> contactIdList, string path)
        {
            List<ContactSituation> ContactSituationList = new List<ContactSituation>();
            try
            {
                foreach (string contactId in contactIdList)
                {
                    ContactSituation contactSituation = new ContactSituation();
                    DataSet ds = Talent.GetInstance().SelectContactSituationDataById(contactId);
                    contactSituation.Info = ds.Tables[0].DBNullToEmpty().DataTableToList<ContactInfo>()[0];
                    contactSituation.Status = ds.Tables[1].DataTableToList<ContactStatus>();
                    string codeList = string.Empty;
                    foreach (DataRow dr in ds.Tables[2].Rows)
                    {
                        codeList += dr[0].ToString() + "\n";
                    }

                    contactSituation.Code = codeList;
                    ContactSituationList.Add(contactSituation);
                }

                string msg = ExcelHelper.GetInstance().ExportMultipleContactSituation(ContactSituationList, path);
                if (msg != "匯出成功")
                {
                    return "匯出失敗";
                }

                return "匯出成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "匯出失敗";
            }
        }

        /// <summary>
        /// 根據聯繫ID匯出面談資料
        /// </summary>
        /// <param name="contactIdList"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        public string ExportInterviewDataByContactId(List<string> contactIdList, string path)
        {
            bool isAllExport = true; ////紀錄是否有資料沒有面談資料
            try
            {
                for (int i = 0; i < contactIdList.Count; i++)
                {
                    List<InterviewData> interviewDataList = new List<InterviewData>(); ///面談資料清單
                    DataSet ds = Talent.GetInstance().SelectContactSituationDataById(contactIdList[i]);
                    List<string> interviewIdList = new List<string>();
                    foreach (DataRow dr in ds.Tables[4].Rows)
                    {
                        interviewIdList.Add(dr[0].ToString());
                    }

                    ////代表沒有面談資料
                    if (interviewIdList.Count == 0)
                    {
                        isAllExport = false;
                        continue;
                    }

                    foreach (string interviewId in interviewIdList)
                    {
                        InterviewResults interviewResults = new InterviewResults();
                        InterviewData interviewData = new InterviewData();
                        DataSet interviewDataSet = Talent.GetInstance().SelectInterviewDataById(interviewId);
                        List<InterviewInfo> interviewInfo = interviewDataSet.Tables[0].DataTableToList<InterviewInfo>(); ////面談基本資料
                        List<InterviewComments> interviewCommentsList = interviewDataSet.Tables[1].DataTableToList<InterviewComments>(); ////面談評語
                        List<InterviewResult> interviewResult = interviewDataSet.Tables[2].DataTableToList<InterviewResult>(); ////面談結果(任用評定與備註)

                        interviewResults.InterviewResult = interviewResult[0];
                        interviewResults.InterviewCommentsList = interviewCommentsList;

                        List<ProjectExperience> projectExperienceList = interviewDataSet.Tables[3].DataTableToList<ProjectExperience>(); ////專案經驗

                        interviewData.InterviewInfo = interviewInfo[0];
                        interviewData.InterviewResults = interviewResults;
                        interviewData.ProjectExperienceList = projectExperienceList;
                        interviewDataList.Add(interviewData);
                    }

                    string msg = ExcelHelper.GetInstance().ExportInterviewData(interviewDataList, path, i + 1);
                    if (msg != "匯出成功")
                    {
                        return "匯出失敗";
                    }
                }

                if (!isAllExport)
                {
                    return "匯出成功，但有些資料沒有面談資料";
                }

                return "匯出成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "匯出失敗";
            }
        }

        /// <summary>
        /// 根據聯繫ID匯出所有資料
        /// </summary>
        /// <param name="contactIdList"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        public string ExportAllDataByContactId(List<string> contactIdList, string path)
        {
            try
            {
                for (int i = 0; i < contactIdList.Count; i++)
                {
                    List<ContactSituation> ContactSituationList = new List<ContactSituation>(); ////聯繫狀況資料
                    List<InterviewData> interviewDataList = new List<InterviewData>(); ///面談資料清單
                    ContactSituation contactSituation = new ContactSituation();
                    DataSet ds = Talent.GetInstance().SelectContactSituationDataById(contactIdList[i]);
                    contactSituation.Info = ds.Tables[0].DBNullToEmpty().DataTableToList<ContactInfo>()[0];
                    contactSituation.Status = ds.Tables[1].DataTableToList<ContactStatus>();
                    string codeList = string.Empty;
                    foreach (DataRow dr in ds.Tables[2].Rows)
                    {
                        codeList += dr[0].ToString() + "\n";
                    }

                    contactSituation.Code = codeList;
                    ContactSituationList.Add(contactSituation);
                    List<string> interviewIdList = new List<string>();
                    foreach (DataRow dr in ds.Tables[4].Rows)
                    {
                        interviewIdList.Add(dr[0].ToString());
                    }

                    foreach (string interviewId in interviewIdList)
                    {
                        InterviewResults interviewResults = new InterviewResults();
                        InterviewData interviewData = new InterviewData();
                        DataSet interviewDataSet = Talent.GetInstance().SelectInterviewDataById(interviewId);
                        List<InterviewInfo> interviewInfo = interviewDataSet.Tables[0].DataTableToList<InterviewInfo>(); ////面談基本資料
                        List<InterviewComments> interviewCommentsList = interviewDataSet.Tables[1].DataTableToList<InterviewComments>(); ////面談評語
                        List<InterviewResult> interviewResult = interviewDataSet.Tables[2].DataTableToList<InterviewResult>(); ////面談結果(任用評定與備註)

                        interviewResults.InterviewResult = interviewResult[0];
                        interviewResults.InterviewCommentsList = interviewCommentsList;

                        List<ProjectExperience> projectExperienceList = interviewDataSet.Tables[3].DataTableToList<ProjectExperience>(); ////專案經驗

                        interviewData.InterviewInfo = interviewInfo[0];
                        interviewData.InterviewResults = interviewResults;
                        interviewData.ProjectExperienceList = projectExperienceList;
                        interviewDataList.Add(interviewData);
                    }

                    string msg = ExcelHelper.GetInstance().ExportAllData(ContactSituationList, interviewDataList, path, i + 1);
                    if (msg != "匯出成功")
                    {
                        return "匯出失敗";
                    }
                }

                return "匯出成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "匯出失敗";
            }
        }

        /// <summary>
        /// 寄送重製密碼的信件
        /// </summary>
        /// <param name="account">欲提醒更換密碼的信箱</param>
        /// <returns>寄送成功or失敗</returns>
        public string AlertUpdatePassword(string account)
        {
            string newPassword = Common.GetInstance().GetRandomPassword(6);
            if (Talent.GetInstance().UpdatePasswordByaccount(account, newPassword) != "修改成功")
            {
                return "密碼重設失敗";
            }

            string mailFrom = ConfigurationManager.AppSettings["MailAccount"];
            string mailPwd = ConfigurationManager.AppSettings["MailPassword"];
            string mailTo = account;
            string mailCc = string.Empty;
            string mailBcc = string.Empty;
            string mailSub = "[人才資料庫] 重設密碼通知";
            string mailBody = "<html><head></head><body><p>Dears</p><p>重設後的密碼如下：</p><p>" + newPassword + "</p></body></html>";
            string msg = Common.GetInstance().SendMail(mailFrom, mailPwd, "smtp.gmail.com", mailTo, mailCc, mailBcc, mailSub, mailBody);
            return msg;
        }

        /// <summary>
        /// 如果值為"不限"，NULL，空，則回傳DBNULL
        /// </summary>
        /// <param name="value">要判斷的值</param>
        /// <returns>會傳值或者是DBNULL</returns>
        public object ValueIsAny(string value)
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

        /// <summary>
        /// 如果帳號沒有Mail格式則幫其補上
        /// </summary>
        /// <param name="account">帳號</param>
        /// <returns>Mail格式的帳號</returns>
        public string AddMailFormat(string account)
        {
            account = (account.Contains("@")) ? account : account + "@is-land.com.tw";
            return account;
        }
    }
}
