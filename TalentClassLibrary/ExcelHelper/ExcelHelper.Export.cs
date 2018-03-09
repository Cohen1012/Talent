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
        /// 匯出多筆聯繫狀況資料
        /// </summary>
        /// <param name="ContactSituationList"></param>
        public string ExportMultipleContactSituation(List<ContactSituation> ContactSituationList, string path)
        {
            try
            {
                //建立Workbook
                Workbook workbook = new Workbook();
                workbook.LoadTemplateFromFile(@".\Template\TalentTemplate.xlsx");

                for (int i = 0; i < ContactSituationList.Count; i++)
                {
                    Worksheet sheet = workbook.CreateEmptySheet();
                    ////第一個Sheet當作Template
                    sheet.CopyFrom(workbook.Worksheets[0]);
                    ////Sheet命名
                    if (!string.IsNullOrEmpty(ContactSituationList[i].Info.Name))
                    {
                        sheet.Name = (i + 1) + "." + ContactSituationList[i].Info.Name;
                    }
                    else if (!string.IsNullOrEmpty(ContactSituationList[i].Code))
                    {
                        string[] code = ContactSituationList[i].Code.Split(new string[] { "\n" }, StringSplitOptions.None);
                        sheet.Name = (i + 1) + "." + code[0];
                    }

                    sheet = CreateContactSituationSheet(ContactSituationList, sheet, i);
                }

                ////刪除Template Sheet
                workbook.Worksheets.Remove("聯繫狀況");
                workbook.Worksheets.Remove("人事資料");
                workbook.Worksheets.Remove("專案經驗");
                workbook.Worksheets.Remove("面談結果");
                //儲存到實體路徑
                string strFullName = Path.Combine(path, "聯繫狀況" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");

                workbook.SaveToFile(strFullName, ExcelVersion.Version2010);
                return "匯出成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "匯出失敗";
            }
        }

        /// <summary>
        /// 匯出面談資料
        /// </summary>
        /// <param name="interviewDataList">面談資料</param>
        /// <param name="path">存檔路徑</param>
        /// <param name="count"></param>
        /// <returns></returns>
        public string ExportInterviewData(List<InterviewData> interviewDataList, string path, int count)
        {
            try
            {
                //建立Workbook
                Workbook workbook = new Workbook();
                workbook.LoadTemplateFromFile(@".\Template\TalentTemplate.xlsx");
                for (int i = 0; i < interviewDataList.Count; i++)
                {
                    ////面談基本資訊Template
                    Worksheet sheet1 = workbook.CreateEmptySheet();
                    sheet1.Name = "人事資料" + (i + 1);
                    sheet1.CopyFrom(workbook.Worksheets[1]);
                    sheet1 = CreateInterviewInfoSheet(interviewDataList[i].InterviewInfo, sheet1);
                    ////專案經驗Template
                    Worksheet sheet2 = workbook.CreateEmptySheet();
                    sheet2.Name = "專案經驗" + (i + 1);
                    sheet2.CopyFrom(workbook.Worksheets[2]);
                    sheet2 = CreateProjectExperienceSheet(interviewDataList[i].ProjectExperienceList, sheet2);
                    ////面談結果Template
                    Worksheet sheet3 = workbook.CreateEmptySheet();
                    sheet3.Name = "面談結果" + (i + 1);
                    sheet3.CopyFrom(workbook.Worksheets[3]);
                    sheet3 = CreateInterviewResualtSheet(interviewDataList[i].InterviewResults, sheet3);
                }

                ////隱藏Template Sheet
                for (int i = 0; i < 4; i++)
                {
                    workbook.Worksheets[i].Visibility = WorksheetVisibility.Hidden;
                }
                //儲存到實體路徑
                string strFullName = Path.Combine(path, count + ".面談資料" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");

                workbook.SaveToFile(strFullName, ExcelVersion.Version2010);
                return "匯出成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "匯出失敗";
            }
        }

        /// <summary>
        /// 匯出所有資料
        /// </summary>
        /// <param name="contactSituationList">聯繫狀況資料</param>
        /// <param name="interviewDataList">面談資料</param>
        /// <param name="path">存檔路徑</param>
        /// <param name="count"></param>
        /// <returns></returns>
        public string ExportAllData(List<ContactSituation> contactSituationList, List<InterviewData> interviewDataList, string path, int count)
        {
            try
            {
                //建立Workbook
                Workbook workbook = new Workbook();
                workbook.LoadTemplateFromFile(@".\Template\TalentTemplate.xlsx");
                //聯繫狀況Template
                Worksheet sheet = workbook.Worksheets[0];
                sheet = CreateContactSituationSheet(contactSituationList, sheet, 0);
                for (int i = 0; i < interviewDataList.Count; i++)
                {
                    ////面談基本資訊Template
                    Worksheet sheet1 = workbook.CreateEmptySheet();
                    sheet1.Name = "人事資料" + (i + 1);
                    sheet1.CopyFrom(workbook.Worksheets[1]);
                    sheet1 = CreateInterviewInfoSheet(interviewDataList[i].InterviewInfo, sheet1);
                    ////專案經驗Template
                    Worksheet sheet2 = workbook.CreateEmptySheet();
                    sheet2.Name = "專案經驗" + (i + 1);
                    sheet2.CopyFrom(workbook.Worksheets[2]);
                    sheet2 = CreateProjectExperienceSheet(interviewDataList[i].ProjectExperienceList, sheet2);
                    ////面談結果Template
                    Worksheet sheet3 = workbook.CreateEmptySheet();
                    sheet3.Name = "面談結果" + (i + 1);
                    sheet3.CopyFrom(workbook.Worksheets[3]);
                    sheet3 = CreateInterviewResualtSheet(interviewDataList[i].InterviewResults, sheet3);
                }

                ////隱藏Template Sheet
                for (int i = 1; i < 4; i++)
                {
                    workbook.Worksheets[i].Visibility = WorksheetVisibility.Hidden;
                }
                //儲存到實體路徑
                string strFullName = Path.Combine(path, count + ".聯繫狀況與面談資料" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");

                workbook.SaveToFile(strFullName, ExcelVersion.Version2010);
                return "匯出成功";
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return "匯出失敗";
            }
        }
    }
}
