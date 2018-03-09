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
        private static ExcelHelper excelHelper = new ExcelHelper();

        public static ExcelHelper GetInstance() => excelHelper;

        /// <summary>
        /// 紀錄大略發生的錯誤
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 讀取面談結果的Sheet
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private DataTable[] ReadInterviewResultSheet(Worksheet sheet)
        {
            DataTable[] dt = new DataTable[2];
            int row = 2; ////紀錄目前在哪一行
            InterviewResult interviewResult = new InterviewResult();
            List<InterviewComments> interviewCommentList = new List<InterviewComments>();
            if (sheet.Range["B" + row].Value.Trim() != "任用評定")
            {
                throw new Exception("面談結果sheet格式不符合");
            }

            row = 5;
            interviewResult.Appointment = string.Empty;
            ////任用評定
            for (int i = 5; i <= 7; i++)
            {
                if (sheet.Range["B" + row].Value.Trim().StartsWith("■"))
                {
                    interviewResult.Appointment = sheet.Range["B" + row].Value.Trim().RemoveStartsWithDelimiter("■");
                    break;
                }
                row++;
            }

            row = 13;
            ////面談評語
            while (sheet.Range["B" + row].Value.Trim() != "備註")
            {
                InterviewComments interviewComments = new InterviewComments
                {
                    Interviewer = sheet.Range["C" + row].Value.Trim(),
                    Result = sheet.Range["G" + row].Value.Trim()
                };

                if (!interviewComments.TResultIsEmtpty())
                {
                    interviewCommentList.Add(interviewComments);
                }
                row++;
                if (sheet.Range["B" + row].Value.Trim() == "備註")
                {
                    break;
                }

                row++;
            }

            row += 3;
            interviewResult.Results_Remark = sheet.Range["B" + row].Value.Trim();
            List<InterviewResult> interviewResultList = new List<InterviewResult>
            {
                interviewResult
            };

            if(interviewCommentList.Count == 0)
            {
                interviewCommentList.Add(new InterviewComments());
            }

            dt[1] = interviewResultList.ListToDataTable();
            dt[0] = interviewCommentList.ListToDataTable();

            return dt;
        }

        /// <summary>
        /// 讀取專案經驗的Sheet
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private DataTable ReadProjectExperienceSheet(Worksheet sheet)
        {
            int row = 2; ////紀錄目前在哪一行
            List<ProjectExperience> projectExperienceList = new List<ProjectExperience>();
            while (!string.IsNullOrEmpty(sheet.Range["B" + row].Value.Trim()) && sheet.Range["B" + row].Value.Trim().Equals("公司名稱"))
            {
                ProjectExperience projectExperience = new ProjectExperience
                {
                    Company = sheet.Range["C" + row].Value.Trim(),
                    Position = sheet.Range["H" + row].Value.Trim(),
                };

                row++;
                projectExperience.Project_Name = sheet.Range["C" + row].Value.Trim();
                projectExperience.Start_End_Date = sheet.Range["H" + row].Value.Trim();
                row++;
                projectExperience.OS = sheet.Range["C" + row].Value.Trim();
                projectExperience.Language = sheet.Range["H" + row].Value.Trim();
                row++;
                projectExperience.Database = sheet.Range["C" + row].Value.Trim();
                projectExperience.Tools = sheet.Range["H" + row].Value.Trim();
                row++;
                projectExperience.Description = sheet.Range["C" + row].Value.Trim();
                row += 3;
                projectExperienceList.Add(projectExperience);
            }

            if(projectExperienceList.Count == 0)
            {
                projectExperienceList.Add(new ProjectExperience());
            }

            return projectExperienceList.ListToDataTable();
        }

        /// <summary>
        /// 讀取面談基本資料的Sheet
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private DataTable ReadInterviewInfoSheet(Worksheet sheet)
        {
            int row = 0; ////紀錄目前在哪一行
            DataTable dt = new DataTable();
            if (sheet.Range["E2"].Value.Trim() != "應徵職缺")
            {
                throw new Exception("面談結果sheet格式不符合");
            }

            InterviewInfo interviewInfo = new InterviewInfo
            {
                Vacancies = sheet.Range["F2"].Value.Trim(),
                Name = sheet.Range["F4"].Value.Trim(),
                Married = sheet.Range["F6"].Value.Trim(),
                Adress = sheet.Range["F8"].Value.Trim(),
                Urgent_Contact_Person = sheet.Range["F10"].Value.Trim(),
                Interview_Date = sheet.Range["I2"].Value.Trim(),
                Sex = sheet.Range["I4"].Value.Trim(),
                Mail = sheet.Range["I6"].Value.Trim(),
                Urgent_Relationship = sheet.Range["I10"].Value.Trim(),
                Birthday = sheet.Range["L4"].Value.Trim(),
                CellPhone = sheet.Range["L6"].Value.Trim(),
                Urgent_CellPhone = sheet.Range["L10"].Value.Trim(),
            };

            ////處理圖片
            if (sheet.Pictures.Count > 0)
            {
                ExcelPicture picture = sheet.Pictures[0];
                string tempPath = Path.GetTempFileName();
                picture.Picture.Save(tempPath, ImageFormat.Png);
                interviewInfo.Image = tempPath;
                //picture.Picture.Save(@"..\..\..\Template\image.png", ImageFormat.Png);
            }

            row = 13;
            ////教育程度
            List<Education> educationList = new List<Education>();
            while (sheet.Range["C" + row].Value.Trim() != "經歷")
            {
                Education education = new Education
                {
                    School = sheet.Range["D" + row].Value.Trim(),
                    Department = sheet.Range["E" + row].Value.Trim(),
                    Start_End_Date = sheet.Range["F" + row].Value.Trim(),
                    Is_Graduation = sheet.Range["G" + row].Value.Trim(),
                    Remark = sheet.Range["H" + row].Value.Trim(),
                };

                if (!education.TResultIsEmtpty())
                {
                    educationList.Add(education);
                }

                row++;
            }
            interviewInfo.Education = JsonConvert.SerializeObject(educationList, Formatting.Indented);
            interviewInfo.Education = interviewInfo.Education == "[]" ? string.Empty : interviewInfo.Education;

            row++;
            ////工作經驗
            List<WorkExperience> workExperienceList = new List<WorkExperience>();
            while (sheet.Range["C" + row].Value.Trim() != "兵役")
            {
                WorkExperience workExperience = new WorkExperience
                {
                    Institution_name = sheet.Range["D" + row].Value.Trim(),
                    Position = sheet.Range["E" + row].Value.Trim(),
                    Start_End_Date = sheet.Range["F" + row].Value.Trim(),
                    Start_Salary = sheet.Range["G" + row].Value.Trim(),
                    End_Salary = sheet.Range["H" + row].Value.Trim(),
                    Leaving_Reason = sheet.Range["I" + row].Value.Trim(),
                };

                if (!workExperience.TResultIsEmtpty())
                {
                    workExperienceList.Add(workExperience);
                }
                row++;
            }
            interviewInfo.Work_Experience = JsonConvert.SerializeObject(workExperienceList, Formatting.Indented);
            interviewInfo.Work_Experience = interviewInfo.Work_Experience == "[]" ? string.Empty : interviewInfo.Work_Experience;

            ////兵役
            row++;
            interviewInfo.IsService = sheet.Range["D" + row].Value.Trim();
            interviewInfo.Exemption_Reason = sheet.Range["E" + row].Value.Trim();
            row += 3;
            ////專長
            ////專長-程式語言
            interviewInfo.Expertise_Language = GetExpertise(sheet, row, 11, 13);
            row += 2;
            ////專長-開發工具
            interviewInfo.Expertise_Tools = GetExpertise(sheet, row, 7, 12);
            interviewInfo.Expertise_Tools_Framwork = sheet.Range[row, 9].Value.Trim();
            row += 2;
            ////專長-Devops
            interviewInfo.Expertise_Devops = GetExpertise(sheet, row, 9, 11);
            row += 2;
            ////專長-作業系統
            interviewInfo.Expertise_OS = GetExpertise(sheet, row, 9, 11);
            row += 2;
            ////專長-大數據
            interviewInfo.Expertise_BigData = GetExpertise(sheet, row, 7, 9);
            row += 2;
            ////專長-資料庫
            interviewInfo.Expertise_DataBase = GetExpertise(sheet, row, 8, 10);
            row += 2;
            ////專長-專業認證
            interviewInfo.Expertise_Certification = GetExpertise(sheet, row, 9, 11);
            row += 3;
            ////語言能力
            List<Language> languageList = new List<Language>();
            while (sheet.Range["C" + row].Value.Trim() != "最近三年內，是否有計畫繼續就學或出國深造?")
            {
                Language language = new Language
                {
                    Language_Name = sheet.Range["D" + row].Value.Trim(),
                    Listen = sheet.Range["E" + row].Value.Trim(),
                    Speak = sheet.Range["F" + row].Value.Trim(),
                    Read = sheet.Range["G" + row].Value.Trim(),
                    Write = sheet.Range["H" + row].Value.Trim(),
                };

                if (!language.TResultIsEmtpty())
                {
                    languageList.Add(language);
                }
                row++;
            }
            interviewInfo.Language = JsonConvert.SerializeObject(languageList, Formatting.Indented);
            interviewInfo.Language = interviewInfo.Language == "[]" ? string.Empty : interviewInfo.Language;
            row++;
            interviewInfo.IsStudy = sheet.Range["C" + row].Value.Trim();
            row += 2;
            interviewInfo.IsService = sheet.Range["F" + row].Value.Trim();
            interviewInfo.Relatives_Relationship = sheet.Range["J" + row].Value.Trim();
            interviewInfo.Relatives_Name = sheet.Range["N" + row].Value.Trim();
            row += 2;
            interviewInfo.Care_Work = sheet.Range["F" + row].Value.Trim();
            interviewInfo.Hope_Salary = sheet.Range["J" + row].Value.Trim();
            interviewInfo.When_Report = sheet.Range["N" + row].Value.Trim();
            row += 2;
            interviewInfo.Advantage = sheet.Range["D" + row].Value.Trim();
            interviewInfo.Disadvantages = sheet.Range["J" + row].Value.Trim();
            row += 2;
            interviewInfo.Hobby = sheet.Range["D" + row].Value.Trim();
            row += 3;
            interviewInfo.Attract_Reason = sheet.Range["C" + row].Value.Trim();
            row += 3;
            interviewInfo.Future_Goal = sheet.Range["C" + row].Value.Trim();
            row += 3;
            interviewInfo.Hope_Supervisor = sheet.Range["C" + row].Value.Trim();
            row += 3;
            interviewInfo.Hope_Promise = sheet.Range["C" + row].Value.Trim();
            row += 3;
            interviewInfo.Introduction = sheet.Range["C" + row].Value.Trim();

            List<InterviewInfo> interviewInfoList = new List<InterviewInfo>
            {
                interviewInfo
            };

            dt = interviewInfoList.ListToDataTable();

            return dt;

        }

        /// <summary>
        /// 打專長組成字串
        /// </summary>
        /// <param name="row">第幾行</param>
        /// <param name="endExpertiseIndex">最後預設的專長的欄數</param>
        /// <param name="otherExpertiseIndex">其他專長的欄數</param>
        /// <returns></returns>
        private string GetExpertise(Worksheet sheet, int row, int endExpertiseIndex, int otherExpertiseIndex)
        {
            string expertise = string.Empty;
            for (int i = 5; i <= endExpertiseIndex; i++)
            {
                if (sheet.Range[row, i].Value.Trim().StartsWith("■"))
                {
                    expertise += sheet.Range[row, i].Value.Trim().RemoveStartsWithDelimiter("■") + ",";
                }
            }

            expertise += sheet.Range[row, otherExpertiseIndex].Value.Trim();
            expertise = expertise.RemoveEndWithDelimiter(",");

            return expertise;
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
