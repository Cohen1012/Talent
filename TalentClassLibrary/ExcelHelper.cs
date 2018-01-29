﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;
using System.IO;
using TalentClassLibrary.Model;
using ServiceStack.Text;
using ServiceStack;
using Spire.Xls.Core;
using System.Data;
using ShareClassLibrary;
using System.Windows.Forms;

namespace TalentClassLibrary
{
    /// <summary>
    /// 處理Excel的類別
    /// </summary>
    public class ExcelHelper
    {
        private static ExcelHelper excelHelper = new ExcelHelper();

        public static ExcelHelper GetInstance() => excelHelper;

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
                workbook.LoadTemplateFromFile(@"..\..\..\Template\TalentTemplate.xlsx");

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

                ////隱藏Template Sheet
                for (int i = 0; i < 4; i++)
                {
                    workbook.Worksheets[i].Visibility = WorksheetVisibility.Hidden;
                }
                //儲存到物理路徑
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
                workbook.LoadTemplateFromFile(@"..\..\..\Template\TalentTemplate.xlsx");
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
                //儲存到物理路徑
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
                workbook.LoadTemplateFromFile(@"..\..\..\Template\TalentTemplate.xlsx");
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
                //儲存到物理路徑
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

        public void ImportOldTalent(string path)
        {
            List<ContactSituation> contactSituationList = new List<ContactSituation>();
            try
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(path);
                Worksheet sheet = workbook.Worksheets[0];
                DataTable dt = sheet.ExportDataTable();

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ContactSituation contactSituation = new ContactSituation();
                    ContactInfo contactInfo = new ContactInfo();
                    ContactStatus contactStatus = new ContactStatus();
                    ////先不處理聯繫狀況
                    for (int j = 0; j < 29; j++)
                    {
                        if(j==3 || j==4 || j== 5)
                        {
                            continue;
                        }

                        switch (dt.Columns[j].ToString())
                        {
                            case "姓名":
                                contactInfo.Name = dt.Rows[i].ItemArray[j].ToString();
                                break;
                            case "地點":
                                contactInfo.Place = dt.Rows[i].ItemArray[j].ToString();
                                break;
                            case "1111/104代碼":
                                contactSituation.Code = dt.Rows[i].ItemArray[j].ToString();
                                break;
                            case "JAVA":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "JSP":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "Android APP":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "ASP":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "C/C++":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "C#":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "ASP.NET":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "VB.NET":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "VB6":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "HTML":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "Javascript":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "Bootstrap":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "Delphi":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "PHP":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "研替":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "Hadoop":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "ETL":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "R":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "notes":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "UI/UX":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "資料庫":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                            case "Linux":
                                contactInfo.Skill += dt.Rows[i].ItemArray[j].ToString() + ",";
                                break;
                        }
                        for (int z = 3; z <= 5; z++)
                        {
                            switch (dt.Columns[j].ToString())
                            {
                                case "日期":
                                    string msg = Valid.GetInstance().ValidDateFormat(dt.Rows[i].ItemArray[j].ToString());
                                    if (msg != string.Empty)
                                    {
                                        MessageBox.Show(msg);
                                        return;
                                    }

                                    contactStatus.Contact_Date = dt.Rows[i].ItemArray[j].ToString();
                                    break;
                                case "聯絡狀況":
                                    contactStatus.Contact_Status = dt.Rows[i].ItemArray[j].ToString();
                                    break;
                                case "說明":
                                    contactStatus.Remarks = dt.Rows[i].ItemArray[j].ToString();
                                    break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
            }
        }

        /// <summary>
        /// 面談資本資料的Sheet
        /// </summary>
        /// <param name="interviewInfo">面談基本資訊</param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private Worksheet CreateInterviewInfoSheet(InterviewInfo interviewInfo, Worksheet sheet)
        {
            int row = 0; ////紀錄目前在哪一行
            ////照片
            if (!string.IsNullOrEmpty(interviewInfo.Image))
            {
                string picPath = interviewInfo.Image;
                ExcelPicture picture = sheet.Pictures.Add(2, 2, picPath);
                picture.Height = 135;
                picture.Width = 145;
            }

            ////面談基本資料
            sheet.Range["F2"].Text = interviewInfo.Vacancies;
            sheet.Range["F4"].Text = interviewInfo.Name;
            sheet.Range["F6"].Text = interviewInfo.Married;
            sheet.Range["F8"].Text = interviewInfo.Adress;
            sheet.Range["F10"].Text = interviewInfo.Urgent_Contact_Person;
            sheet.Range["I2"].Text = interviewInfo.Interview_Date;
            sheet.Range["I4"].Text = interviewInfo.Sex;
            sheet.Range["I6"].Text = interviewInfo.Mail;
            sheet.Range["I10"].Text = interviewInfo.Urgent_Relationship;
            sheet.Range["L4"].Text = string.IsNullOrEmpty(interviewInfo.Birthday) ? string.Empty : interviewInfo.Birthday;
            sheet.Range["L6"].Text = interviewInfo.CellPhone;
            sheet.Range["L10"].Text = interviewInfo.Urgent_CellPhone;
            ////學歷
            List<Education> educationList = new List<Education>();
            if (!string.IsNullOrEmpty(interviewInfo.Education))
            {
                educationList = interviewInfo.Education.FromJson<List<Education>>();
            }

            row = 12;
            for (int i = 0; i < educationList.Count; i++)
            {
                row++;
                sheet.Range[row, 4].Text = string.IsNullOrEmpty(educationList[i].School) ? string.Empty : educationList[i].School;
                sheet.Range[row, 5].Text = string.IsNullOrEmpty(educationList[i].Department) ? string.Empty : educationList[i].Department;
                sheet.Range[row, 6].Text = string.IsNullOrEmpty(educationList[i].Start_End_Date) ? string.Empty : educationList[i].Start_End_Date;
                sheet.Range[row, 7].Text = string.IsNullOrEmpty(educationList[i].Is_Graduation) ? string.Empty : educationList[i].Is_Graduation;
                sheet.Range[row, 8].Text = string.IsNullOrEmpty(educationList[i].Remark) ? string.Empty : educationList[i].Remark;
                sheet.Range["H" + row + ":M" + row].Merge();
                sheet.Range["D" + row + ":H" + row].Style.HorizontalAlignment = HorizontalAlignType.Left;
            }
            sheet.Range["D12:M" + row].BorderAround(LineStyleType.Medium);
            sheet.Range["D12:M" + row].BorderInside(LineStyleType.Thin);
            /////工作經驗
            List<WorkExperience> workExperienceList = new List<WorkExperience>();
            if (!string.IsNullOrEmpty(interviewInfo.Work_Experience))
            {
                workExperienceList = interviewInfo.Work_Experience.FromJson<List<WorkExperience>>();
            }

            row += 2;
            sheet.Range[row, 3].Text = "經歷";
            sheet.Range[row, 4].Text = "機構名稱";
            sheet.Range[row, 5].Text = "職稱";
            sheet.Range[row, 6].Text = "起迄年月";
            sheet.Range[row, 7].Text = "到職薪資";
            sheet.Range[row, 8].Text = "離職薪資";
            sheet.Range[row, 9].Text = "離職原因";
            sheet.Range["I" + row + ":M" + row].Merge();
            sheet.Range["D" + row + ":I" + row].Style.HorizontalAlignment = HorizontalAlignType.Center;

            for (int i = 0; i < workExperienceList.Count; i++)
            {
                row++;
                sheet.Range[row, 4].Text = string.IsNullOrEmpty(workExperienceList[i].Institution_name) ? string.Empty : workExperienceList[i].Institution_name;
                sheet.Range[row, 5].Text = string.IsNullOrEmpty(workExperienceList[i].Position) ? string.Empty : workExperienceList[i].Position;
                sheet.Range[row, 6].Text = string.IsNullOrEmpty(workExperienceList[i].Start_End_Date) ? string.Empty : workExperienceList[i].Start_End_Date;
                sheet.Range[row, 7].Text = string.IsNullOrEmpty(workExperienceList[i].Start_Salary) ? string.Empty : workExperienceList[i].Start_Salary;
                sheet.Range[row, 8].Text = string.IsNullOrEmpty(workExperienceList[i].End_Salary) ? string.Empty : workExperienceList[i].End_Salary;
                sheet.Range[row, 9].Text = string.IsNullOrEmpty(workExperienceList[i].Leaving_Reason) ? string.Empty : workExperienceList[i].Leaving_Reason;
                sheet.Range["I" + row + ":M" + row].Merge();
                sheet.Range["D" + row + ":M" + row].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["G" + row + ":H" + row].Style.HorizontalAlignment = HorizontalAlignType.Right;
            }
            sheet.Range["D" + (row - workExperienceList.Count) + ":M" + row].BorderAround(LineStyleType.Medium);
            sheet.Range["D" + (row - workExperienceList.Count) + ":M" + row].BorderInside(LineStyleType.Thin);
            ////兵役
            row += 2;
            sheet.Range[row, 3].Text = "兵役";
            sheet.Range[row, 4].Text = "服役期間";
            sheet.Range[row, 5].Text = "免役說明";
            sheet.Range["E" + row + ":M" + row].Merge();
            sheet.Range["D" + row + ":E" + row].Style.HorizontalAlignment = HorizontalAlignType.Center;
            row++;
            sheet.Range[row, 4].Text = interviewInfo.During_Service;
            sheet.Range[row, 5].Text = interviewInfo.Exemption_Reason;
            sheet.Range["E" + row + ":M" + row].Merge();
            sheet.Range["D" + row + ":M" + row].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["D" + (row - 1) + ":M" + row].BorderAround(LineStyleType.Medium);
            sheet.Range["D" + (row - 1) + ":M" + row].BorderInside(LineStyleType.Thin);
            ////專長
            row += 2;
            sheet.Range[row, 3].Text = "專長";
            sheet.Range[row, 4].Text = "※請盡可能寫出您所具備的專長及與工作有關的技能";
            sheet.Range["D" + row + ":G" + row].Merge();
            sheet.Range["D" + row + ":G" + row].Style.HorizontalAlignment = HorizontalAlignType.Left;
            row++;
            ////----程式語言
            sheet.Range[row, 4].Text = "程式語言";
            string[] Expertise_Language = { "C/C++", "Java", "C#", "VB", "Scala", "R", "Python" };
            for (int i = 5; i <= 11; i++)
            {
                sheet.Range[row, i].Text = "□" + Expertise_Language[i - 5];
            }

            sheet.Range[row, 12].Text = "其他";
            sheet.Range["M" + row + ":O" + row].Merge();
            sheet.Range["M" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 13].Style.HorizontalAlignment = HorizontalAlignType.Left;
            ChcekedExpertise(interviewInfo.Expertise_Language, sheet, row, Expertise_Language, 13);

            ////----開發工具
            row += 2;
            sheet.Range[row, 4].Text = "開發工具";
            string[] Expertise_Tools = { "Visual Studio", "Eclipse", "Xcode" };
            for (int i = 5; i <= 7; i++)
            {
                sheet.Range[row, i].Text = "□" + Expertise_Tools[i - 5];
            }

            sheet.Range[row, 8].Text = "Framwork";
            sheet.Range["I" + row + ":J" + row].Merge();
            sheet.Range["I" + row + ":J" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 9].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 11].Text = "其他";
            sheet.Range["L" + row + ":O" + row].Merge();
            sheet.Range["L" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 12].Style.HorizontalAlignment = HorizontalAlignType.Left;
            if (!string.IsNullOrEmpty(interviewInfo.Expertise_Tools))
            {
                string[] expertiseTools = interviewInfo.Expertise_Tools.Split(',');
                for (int i = 0; i < expertiseTools.Length; i++)
                {
                    for (int j = 0; j < Expertise_Tools.Length; j++)
                    {
                        if (Expertise_Tools[j] == expertiseTools[i])
                        {
                            sheet.Range[row, (j + 5)].Text = "■" + Expertise_Tools[j];
                            break;
                        }

                        if (j == Expertise_Tools.Length - 1)
                        {
                            if (expertiseTools[i].StartsWith("Framwork："))
                            {
                                sheet.Range[row, 9].Text += expertiseTools[i].Substring(9);
                            }
                            else
                            {
                                sheet.Range[row, 12].Text += expertiseTools[i] + ",";
                            }
                        }
                    }
                }

                sheet.Range[row, 12].Text = this.RemoveEndWithComma(sheet.Range[row, 12].Text);
            }
            ////----Devops
            row += 2;
            sheet.Range[row, 4].Text = "Devops";
            string[] Expertise_Devops = { "Docker", "Git", "Vagrant", "Puppet", "TFS" };
            for (int i = 5; i <= 9; i++)
            {
                sheet.Range[row, i].Text = "□" + Expertise_Devops[i - 5];
            }

            sheet.Range[row, 10].Text = "其他";
            sheet.Range["K" + row + ":O" + row].Merge();
            sheet.Range["K" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 11].Style.HorizontalAlignment = HorizontalAlignType.Left;
            ChcekedExpertise(interviewInfo.Expertise_Devops, sheet, row, Expertise_Devops, 11);
            ////----作業系統
            row += 2;
            sheet.Range[row, 4].Text = "作業系統";
            string[] Expertise_OS = { "Windows", "Linux", "Unix", "Android", "IOS" };
            for (int i = 5; i <= 9; i++)
            {
                sheet.Range[row, i].Text = "□" + Expertise_OS[i - 5];
            }

            sheet.Range[row, 10].Text = "其他";
            sheet.Range["K" + row + ":O" + row].Merge();
            sheet.Range["K" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 11].Style.HorizontalAlignment = HorizontalAlignType.Left;
            ChcekedExpertise(interviewInfo.Expertise_OS, sheet, row, Expertise_OS, 11);
            ////----大數據
            row += 2;
            sheet.Range[row, 4].Text = "大數據";
            string[] Expertise_BigData = { "Hadoop", "Spark", "Cloudera" };
            for (int i = 5; i <= 7; i++)
            {
                sheet.Range[row, i].Text = "□" + Expertise_BigData[i - 5];
            }

            sheet.Range[row, 8].Text = "其他";
            sheet.Range["I" + row + ":O" + row].Merge();
            sheet.Range["I" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 9].Style.HorizontalAlignment = HorizontalAlignType.Left;
            ChcekedExpertise(interviewInfo.Expertise_BigData, sheet, row, Expertise_BigData, 9);
            ////----資料庫
            row += 2;
            sheet.Range[row, 4].Text = "資料庫";
            string[] Expertise_DataBase = { "Oracle", "My SQL", "MS SQL", "Hbase" };
            for (int i = 5; i <= 8; i++)
            {
                sheet.Range[row, i].Text = "□" + Expertise_DataBase[i - 5];
            }

            sheet.Range[row, 9].Text = "其他";
            sheet.Range["J" + row + ":O" + row].Merge();
            sheet.Range["J" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 10].Style.HorizontalAlignment = HorizontalAlignType.Left;
            ChcekedExpertise(interviewInfo.Expertise_DataBase, sheet, row, Expertise_DataBase, 10);
            ////----專業認證
            row += 2;
            sheet.Range[row, 4].Text = "專業認證";
            string[] Expertise_Certification = { "SCJP", "SCWCD", "MCAD", "MCTS", "PMP" };
            for (int i = 5; i <= 9; i++)
            {
                sheet.Range[row, i].Text = "□" + Expertise_Certification[i - 5];
            }

            sheet.Range[row, 10].Text = "其他";
            sheet.Range["K" + row + ":O" + row].Merge();
            sheet.Range["K" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 11].Style.HorizontalAlignment = HorizontalAlignType.Left;
            ChcekedExpertise(interviewInfo.Expertise_Certification, sheet, row, Expertise_Certification, 11);
            /////語言能力
            List<Language> languageList = new List<Language>();
            if (!string.IsNullOrEmpty(interviewInfo.Language))
            {
                languageList = interviewInfo.Language.FromJson<List<Language>>();
            }

            row += 2;
            sheet.Range[row, 3].Text = "語言能力";
            sheet.Range[row, 4].Text = "語言";
            sheet.Range[row, 5].Text = "聽";
            sheet.Range[row, 6].Text = "說";
            sheet.Range[row, 7].Text = "讀";
            sheet.Range[row, 8].Text = "寫";
            sheet.Range["D" + row + ":H" + row].Style.HorizontalAlignment = HorizontalAlignType.Center;

            for (int i = 0; i < languageList.Count; i++)
            {
                row++;
                sheet.Range[row, 4].Text = string.IsNullOrEmpty(languageList[i].Language_Name) ? string.Empty : languageList[i].Language_Name;
                sheet.Range[row, 5].Text = string.IsNullOrEmpty(languageList[i].Listen) ? string.Empty : languageList[i].Listen;
                sheet.Range[row, 6].Text = string.IsNullOrEmpty(languageList[i].Speak) ? string.Empty : languageList[i].Speak;
                sheet.Range[row, 7].Text = string.IsNullOrEmpty(languageList[i].Read) ? string.Empty : languageList[i].Read;
                sheet.Range[row, 8].Text = string.IsNullOrEmpty(languageList[i].Write) ? string.Empty : languageList[i].Write;
                sheet.Range["D" + row + ":H" + row].Style.HorizontalAlignment = HorizontalAlignType.Center;
            }
            sheet.Range["D" + (row - languageList.Count) + ":H" + row].BorderAround(LineStyleType.Medium);
            sheet.Range["D" + (row - languageList.Count) + ":H" + row].BorderInside(LineStyleType.Thin);
            ////問題
            row += 2;
            sheet.Range[row, 3].Text = "最近三年內，是否有計畫繼續就學或出國深造?";
            sheet.Range["C" + row + ":F" + row].Merge();
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            row++;
            sheet.Range["C" + row + ":O" + row].Merge();
            sheet.Range["C" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 3].Text = interviewInfo.IsStudy;
            row += 2;
            sheet.Range[row, 3].Text = "有無親友在本公司服務?";
            sheet.Range["C" + row + ":E" + row].Merge();
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["F" + row + ":G" + row].Merge();
            sheet.Range["F" + row + ":G" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 6].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 6].Text = interviewInfo.IsService;
            sheet.Range[row, 9].Text = "關係";
            sheet.Range[row, 9].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["J" + row + ":K" + row].Merge();
            sheet.Range["J" + row + ":K" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 10].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 10].Text = interviewInfo.Relatives_Relationship;
            sheet.Range[row, 13].Text = "姓名";
            sheet.Range[row, 13].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["N" + row + ":O" + row].Merge();
            sheet.Range["N" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 14].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 14].Text = interviewInfo.Relatives_Name;
            row += 2;
            sheet.Range[row, 3].Text = "在工作中您最在乎什麼?";
            sheet.Range["C" + row + ":E" + row].Merge();
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["F" + row + ":G" + row].Merge();
            sheet.Range["F" + row + ":G" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 6].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 6].Text = interviewInfo.Care_Work;
            sheet.Range[row, 9].Text = "希望待遇";
            sheet.Range[row, 9].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["J" + row + ":K" + row].Merge();
            sheet.Range["J" + row + ":K" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 10].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 10].Text = interviewInfo.Hope_Salary;
            sheet.Range[row, 13].Text = "通知後多久可報到?";
            sheet.Range[row, 13].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["N" + row + ":O" + row].Merge();
            sheet.Range["N" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 14].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 14].Text = interviewInfo.When_Report;
            row += 2;
            sheet.Range[row, 3].Text = "優點";
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["D" + row + ":G" + row].Merge();
            sheet.Range["D" + row + ":G" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 4].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 4].Text = interviewInfo.Advantage;
            sheet.Range[row, 9].Text = "缺點";
            sheet.Range[row, 9].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["J" + row + ":O" + row].Merge();
            sheet.Range["J" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 10].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 10].Text = interviewInfo.Disadvantages;
            row += 2;
            sheet.Range[row, 3].Text = "嗜好";
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["D" + row + ":G" + row].Merge();
            sheet.Range["D" + row + ":G" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 4].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 4].Text = interviewInfo.Hobby;
            row += 2;
            sheet.Range[row, 3].Text = "什麼原因吸引您來亦思應徵此工作? 您對亦思的第一印象是什麼?";
            sheet.Range["C" + row + ":G" + row].Merge();
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            row++;
            sheet.Range["C" + row + ":O" + row].Merge();
            sheet.Range["C" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 3].Text = interviewInfo.Attract_Reason;
            row += 2;
            sheet.Range[row, 3].Text = "請描述您對未來的目標、希望。";
            sheet.Range["C" + row + ":E" + row].Merge();
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            row++;
            sheet.Range["C" + row + ":O" + row].Merge();
            sheet.Range["C" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 3].Text = interviewInfo.Future_Goal;
            row += 2;
            sheet.Range[row, 3].Text = "您心目中的理想主管是什麼樣的人?";
            sheet.Range["C" + row + ":E" + row].Merge();
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            row++;
            sheet.Range["C" + row + ":O" + row].Merge();
            sheet.Range["C" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 3].Text = interviewInfo.Hope_Supervisor;
            row += 2;
            sheet.Range[row, 3].Text = "您希望亦思給與您什麼承諾?(任何方面)";
            sheet.Range["C" + row + ":E" + row].Merge();
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            row++;
            sheet.Range["C" + row + ":O" + row].Merge();
            sheet.Range["C" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 3].Text = interviewInfo.Hope_Promise;
            row += 2;
            sheet.Range[row, 3].Text = "覺得自己的哪個特質最強烈? 現在您有三十秒的自我推薦時間，您會透過什麼方式讓我們對您印象深刻?";
            sheet.Range["C" + row + ":K" + row].Merge();
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            row++;
            sheet.Range["C" + row + ":O" + row].Merge();
            sheet.Range["C" + row + ":O" + row].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
            sheet.Range[row, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range[row, 3].Text = interviewInfo.Introduction;
            return sheet;
        }

        /// <summary>
        /// 勾選專長
        /// </summary>
        /// <param name="Expertise">專長標題</param>
        /// <param name="sheet"></param>
        /// <param name="row">第幾列</param>
        /// <param name="DefaultExpertise">該專長的預設項目</param>
        /// <param name="column">"其他"欄位所在的Column</param>
        /// <returns></returns>
        private Worksheet ChcekedExpertise(string Expertise, Worksheet sheet, int row, string[] DefaultExpertise, int column)
        {
            if (!string.IsNullOrEmpty(Expertise))
            {
                string[] expertiseDevops = Expertise.Split(',');
                for (int i = 0; i < expertiseDevops.Length; i++)
                {
                    for (int j = 0; j < DefaultExpertise.Length; j++)
                    {
                        if (DefaultExpertise[j] == expertiseDevops[i])
                        {
                            sheet.Range[row, (j + 5)].Text = "■" + DefaultExpertise[j];
                            break;
                        }

                        if (j == DefaultExpertise.Length - 1)
                        {
                            sheet.Range[row, column].Text += expertiseDevops[i] + ",";
                        }
                    }
                }

                sheet.Range[row, column].Text = this.RemoveEndWithComma(sheet.Range[row, column].Text);
            }

            return sheet;
        }

        /// <summary>
        /// 專案經驗的Sheet
        /// </summary>
        /// <param name="projectExperienceList">專案經驗資料</param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private Worksheet CreateProjectExperienceSheet(List<ProjectExperience> projectExperienceList, Worksheet sheet)
        {
            int rowCount = 2; ////紀錄目前所在的列數
            for (int i = 0; i < projectExperienceList.Count; i++)
            {
                sheet.Range[rowCount, 2].Text = "公司名稱";
                sheet.Range[rowCount, 2].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["C" + rowCount + ":E" + rowCount].Merge();
                sheet.Range["C" + rowCount + ":E" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range[rowCount, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range[rowCount, 3].Text = projectExperienceList[i].Company;
                sheet.Range[rowCount, 7].Text = "職稱";
                sheet.Range[rowCount, 7].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["H" + rowCount + ":J" + rowCount].Merge();
                sheet.Range["H" + rowCount + ":J" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range[rowCount, 8].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range[rowCount, 8].Text = projectExperienceList[i].Position;
                rowCount++;
                sheet.Range[rowCount, 2].Text = "專案名稱";
                sheet.Range[rowCount, 2].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["C" + rowCount + ":E" + rowCount].Merge();
                sheet.Range["C" + rowCount + ":E" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range[rowCount, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range[rowCount, 3].Text = projectExperienceList[i].Project_Name;
                sheet.Range[rowCount, 7].Text = "起迄年月";
                sheet.Range[rowCount, 7].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["H" + rowCount + ":J" + rowCount].Merge();
                sheet.Range["H" + rowCount + ":J" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range[rowCount, 8].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range[rowCount, 8].Text = projectExperienceList[i].Start_End_Date;
                rowCount++;
                sheet.Range[rowCount, 2].Text = "作業系統";
                sheet.Range[rowCount, 2].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["C" + rowCount + ":E" + rowCount].Merge();
                sheet.Range["C" + rowCount + ":E" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range[rowCount, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range[rowCount, 3].Text = projectExperienceList[i].OS;
                sheet.Range[rowCount, 7].Text = "程式語言";
                sheet.Range[rowCount, 7].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["H" + rowCount + ":J" + rowCount].Merge();
                sheet.Range["H" + rowCount + ":J" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range[rowCount, 8].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range[rowCount, 8].Text = projectExperienceList[i].Language;
                rowCount++;
                sheet.Range[rowCount, 2].Text = "資料庫";
                sheet.Range[rowCount, 2].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["C" + rowCount + ":E" + rowCount].Merge();
                sheet.Range["C" + rowCount + ":E" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range[rowCount, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range[rowCount, 3].Text = projectExperienceList[i].Database;
                sheet.Range[rowCount, 7].Text = "開發工具";
                sheet.Range[rowCount, 7].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["H" + rowCount + ":J" + rowCount].Merge();
                sheet.Range["H" + rowCount + ":J" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range[rowCount, 8].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range[rowCount, 8].Text = projectExperienceList[i].Tools;
                rowCount++;
                sheet.Range[rowCount, 2].Text = "專案簡述";
                sheet.Range[rowCount, 2].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["C" + rowCount + ":L" + rowCount].Merge();
                sheet.Range["C" + rowCount + ":L" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range[rowCount, 3].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range[rowCount, 3].Text = projectExperienceList[i].Description;
                rowCount++;
                sheet.Range["B" + (rowCount - 5) + ":M" + rowCount].BorderAround(LineStyleType.Thick);
                rowCount += 2;
            }
            return sheet;
        }

        /// <summary>
        /// 面談結果
        /// </summary>
        /// <param name="interviewResults"></param>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private Worksheet CreateInterviewResualtSheet(InterviewResults interviewResults, Worksheet sheet)
        {
            int rowCount = 11;
            ////任用評定
            InterviewResult interviewResult = interviewResults.InterviewResult;
            switch (interviewResult.Appointment)
            {
                case "錄用":
                    sheet.Range["B5"].Text = "■" + interviewResult.Appointment;
                    break;
                case "不錄用":
                    sheet.Range["B6"].Text = "■" + interviewResult.Appointment;
                    break;
                case "暫保留":
                    sheet.Range["B7"].Text = "■" + interviewResult.Appointment;
                    break;
            }
            ////面談評語
            List<InterviewComments> interviewCommentsList = interviewResults.InterviewCommentsList;
            if (interviewCommentsList.Count == 0)
            {
                rowCount += 2;
            }

            for (int i = 0; i < interviewCommentsList.Count; i++)
            {
                rowCount += 2;
                sheet.Range["B" + rowCount].Text = "面談者";
                sheet.Range["B" + rowCount].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["C" + rowCount + ":D" + rowCount].Merge();
                sheet.Range["C" + rowCount + ":D" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range["C" + rowCount].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["C" + rowCount].Text = interviewCommentsList[i].Interviewer;
                sheet.Range["F" + rowCount].Text = "面談結果";
                sheet.Range["F" + rowCount].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["G" + rowCount + ":L" + rowCount].Merge();
                sheet.Range["G" + rowCount + ":L" + rowCount].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Double;
                sheet.Range["G" + rowCount].Style.HorizontalAlignment = HorizontalAlignType.Left;
                sheet.Range["G" + rowCount].Text = interviewCommentsList[i].Result;
            }
            rowCount++;
            sheet.Range["B" + (rowCount - interviewCommentsList.Count - 2) + ":M" + rowCount].BorderAround(LineStyleType.Thin);
            rowCount += 2;
            sheet.Range["B" + rowCount + ":C" + (rowCount + 1)].Merge();
            sheet.Range["B" + rowCount].Text = "備註";
            sheet.Range["B" + rowCount + ":C" + (rowCount + 1)].BorderAround(LineStyleType.Thin);
            rowCount += 3;
            sheet.Range["B" + rowCount + ":M" + (rowCount + 2)].Merge();
            sheet.Range["B" + rowCount].Style.HorizontalAlignment = HorizontalAlignType.Left;
            sheet.Range["B" + rowCount].Text = string.IsNullOrEmpty(interviewResult.Results_Remark) ? string.Empty : interviewResult.Results_Remark;
            sheet.Range["B" + rowCount + ":M" + (rowCount + 2)].BorderAround(LineStyleType.Thin);


            return sheet;
        }
        /// <summary>
        /// 聯繫狀況的Sheet
        /// </summary>
        /// <param name="ContactSituationList">聯繫狀況資料</param>
        /// <param name="sheet"></param>
        /// <param name="i"></param>
        /// <returns></returns>
        private Worksheet CreateContactSituationSheet(List<ContactSituation> ContactSituationList, Worksheet sheet, int i)
        {
            sheet.Range["C2"].Text = ContactSituationList[i].Info.Name;
            sheet.Range["C2"].Style.Font.Underline = FontUnderlineType.None;
            sheet.Range["C4"].Text = ContactSituationList[i].Code;
            sheet.Range["C4"].Style.WrapText = true;
            sheet.Range["C6"].Text = ContactSituationList[i].Info.Sex;
            sheet.Range["C8"].Text = ContactSituationList[i].Info.Mail;
            sheet.Range["C10"].Text = ContactSituationList[i].Info.CellPhone;

            sheet.Range["J2"].Text = ContactSituationList[i].Info.Place;
            sheet.Range["J4"].Text = ContactSituationList[i].Info.Skill;
            sheet.Range["J6"].Text = ContactSituationList[i].Info.Year;
            sheet.Range["J8"].Text = ContactSituationList[i].Info.Cooperation_Mode;
            sheet.Range["J10"].Text = ContactSituationList[i].Info.Status;

            List<ContactStatus> contactStatusList = ContactSituationList[i].Status;
            for (int j = 0; j < contactStatusList.Count; j++)
            {
                sheet.Range[j + 13, 2].Text = contactStatusList[j].Contact_Date;
                var mergecolumn = sheet.Merge(sheet.Range[("C" + (13 + j))], sheet.Range["D" + (13 + j)]);
                mergecolumn.HorizontalAlignment = HorizontalAlignType.Left;
                mergecolumn.Text = contactStatusList[j].Contact_Status;
                sheet.Range["E" + (13 + j) + ":K" + (13 + j)].Merge();
                sheet.Range[j + 13, 5].Text = contactStatusList[j].Remarks;
            }

            sheet.Range["B12:K" + (12 + contactStatusList.Count)].BorderAround(LineStyleType.Medium);
            sheet.Range["B12:K" + (12 + contactStatusList.Count)].BorderInside(LineStyleType.Thin);
            return sheet;
        }

        /// <summary>
        /// 如果字串最後一個字為","則將它移除
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private string RemoveEndWithComma(string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return string.Empty;
            }

            return str.EndsWith(",") ? str.Remove(str.Length - 1) : str;
        }
    }
}
