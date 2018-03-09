using Microsoft.VisualStudio.TestTools.UnitTesting;
using TalentClassLibrary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using ShareClassLibrary;

namespace TalentClassLibrary.Tests
{
    [TestClass()]
    public class TalentTests
    {
        [TestMethod()]
        public void AlertUpdatePasswordTest()
        {
            
            var input = "monkey60146@yahoo.com.tw";
            var expect = "寄送成功";
            var actual = Talent.GetInstance().AlertUpdatePassword(input);
            Assert.AreEqual(expect, actual);
        }

        [TestMethod()]
        public void InsertContactSituationTest()
        {
            var input = TestData.GetInstance().TestContactSituationData();
            var expect = "新增成功";
            var actual = Talent.GetInstance().InsertData(input, "2", "Contact_Situation");
            Assert.AreEqual(expect, actual);
        }

        [TestMethod()]
        public void InsertContactSituationInfoDataTest()
        {
            ////新增聯繫狀況
            var input = TestData.GetInstance().TestContactSituationInfoData();
            var expect = "新增成功";
            var actual = Talent.GetInstance().InsertContactSituationInfoData(input);
            Assert.AreEqual(expect, actual);
        }

        [TestMethod()]
        public void InsertMemberTest()
        {
            var input = "hrTest@is-land.com.tw";
            var expect = "新增成功";
            var actual = Talent.GetInstance().InsertMember(input);
            Assert.AreEqual(expect, actual);
            ////使用非本公司信箱，則不給新增
            var input1 = "hrTest@gmail.con";
            var expect1 = "帳號須為本公司之帳號";
            var actual1 = Talent.GetInstance().InsertMember(input1);
            Assert.AreEqual(expect1, actual1);
        }

        [TestMethod()]
        public void SaveProjectExperienceTest()
        {
            var input = TestData.GetInstance().TestProjectExperienceData();
            var expect = "儲存成功";
            var actual = Talent.GetInstance().SaveProjectExperience(input, "4");
            Assert.AreEqual(expect, actual);
        }

        [TestMethod()]
        public void SelectIdByFilterTest()
        {
            var expected = new List<string> { "2" };
            var actual = Talent.GetInstance().SelectIdByFilter(string.Empty, string.Empty, string.Empty, "合約", "不限", string.Empty, string.Empty, "不限", "不限", string.Empty, string.Empty);
            Assert.AreEqual(true, expected.SequenceEqual(actual));
        }

        [TestMethod()]
        public void SelectContactSituationInfoByCodeTest()
        {
            var input = new List<string>() { "0203", "0204" };
            var expected = 2;
            var actual = Talent.GetInstance().SelectContactSituationInfoByCode(input).Rows.Count;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void SelectMemberInfoTest()
        {
            var expected = true;
            var actual = Talent.GetInstance().SelectMemberInfo().Rows.Count > 0 ? true : false;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void SelectMemberInfoByAccountTest()
        {
            var input = "hr@is-land.com.tw";
            var expected = 1;
            var actual = Talent.GetInstance().SelectMemberInfoByAccount(input).Rows.Count;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void SigInTest()
        {
            var expected = "hr@is-land.com.tw";
            var actual = Talent.GetInstance().SignIn("hr@is-land.com.tw", "hr");
            Assert.AreEqual(expected, actual);
            ////密碼大寫也可以通過
            var expected4 = "hr@is-land.com.tw";
            var actual4 = Talent.GetInstance().SignIn("hr@is-land.com.tw", "HR");
            Assert.AreEqual(expected4, actual4);

            var expected1 = "登入失敗";
            var actual1 = Talent.GetInstance().SignIn("hrQQ@is-land.com.tw", "hrQQ");
            Assert.AreEqual(expected1, actual1);

            var expected2 = "該帳號停用中";
            var actual2 = Talent.GetInstance().SignIn("hrDisabled@is-land.com.tw", "hrDisabled");
            Assert.AreEqual(expected2, actual2);

            var expected3 = "hr@is-land.com.tw";
            var actual3 = Talent.GetInstance().SignIn("hr", "Hr");
            Assert.AreEqual(expected3, actual3);
        }

        [TestMethod()]
        public void UpdateContactSituationInfoDataTest()
        {
            var input = TestData.GetInstance().TestContactSituationInfoData();
            var expect = "修改成功";
            var actual = Talent.GetInstance().UpdateContactSituationInfoData(input, "2");
            Assert.AreEqual(expect, actual);
        }

        [TestMethod()]
        public void UpdatePasswordByaccountTest()
        {
            ////account=hr@is-land.com.tw,oldPassword=abcdefg,newPassword=1234567,newCheckPassword=1234567
            var expected = "修改成功";
            var actual = Talent.GetInstance().UpdatePasswordByaccount("hrTest@is-land.com.tw", "hrTest", "hrTest", "hrTest");
            Assert.AreEqual(expected, actual);
            ////account=hr@is-land.com.tw,oldPassword=abcdeg,newPassword=1234567,newCheckPassword=1234567
            var expected1 = "密碼錯誤";
            var actual1 = Talent.GetInstance().UpdatePasswordByaccount("hrTest@is-land.com.tw", "abcdeg", "1234567", "1234567");
            Assert.AreEqual(expected1, actual1);
            ////account=hr@is-land.com.tw,oldPassword=abcdefg,newPassword=12345678,newCheckPassword=1234567
            var expected2 = "新密碼不一致";
            var actual2 = Talent.GetInstance().UpdatePasswordByaccount("hrTest@is-land.com.tw", "hrTest", "12345678", "1234567");
            Assert.AreEqual(expected2, actual2);
        }

        [TestMethod()]
        public void ValidContactSituationDataTest()
        {
            var input = TestData.GetInstance().TestContactSituationData();
            var expected = string.Empty;
            var actual = Talent.GetInstance().ValidContactSituationData(input);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void ValidContactSituationInfoDataTest()
        {
            var expected = string.Empty;
            var actual = Talent.GetInstance().ValidContactSituationInfoData("賴志維", TestData.GetInstance().TestCodeData(), "男", "monkey60146@yahoo.com.tw", "0978535528", "雲林縣", "MS SQL,Asp.Net,C#", "全職", "追蹤");
            Assert.AreEqual(expected, actual);

            var expected1 = string.Empty;
            var actual1 = Talent.GetInstance().ValidContactSituationInfoData("賴志維", new DataTable(), string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, "全職", string.Empty);
            Assert.AreEqual(expected1, actual1);

            var expected2 = "姓名或代碼請至少填一個\n";
            var actual2 = Talent.GetInstance().ValidContactSituationInfoData(string.Empty, new DataTable(), string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, "皆可", string.Empty);
            Assert.AreEqual(expected2, actual2);
        }

        [TestMethod()]
        public void ValidContactStatusTest()
        {
            var input = "104邀約";
            var expected = string.Empty;
            var actual = Talent.GetInstance().ValidContactStatus(input);
            Assert.AreEqual(expected, actual);

            var input1 = "同意邀約";
            var expected1 = string.Empty;
            var actual1 = Talent.GetInstance().ValidContactStatus(input1);
            Assert.AreEqual(expected1, actual1);

            var input2 = "單純聊天";
            var expected2 = "沒有\"" + input2 + "\"此聯繫狀況";
            var actual2 = Talent.GetInstance().ValidContactStatus(input2);
            Assert.AreEqual(expected2, actual2);
        }

        [TestMethod()]
        public void ValidCooperationModeTest()
        {
            ////全職
            var input = "全職";
            var expected = string.Empty;
            var actual = Talent.GetInstance().ValidCooperationMode(input);
            Assert.AreEqual(expected, actual);
            ////合約
            var input1 = "合約";
            var expected1 = string.Empty;
            var actual1 = Talent.GetInstance().ValidCooperationMode(input1);
            Assert.AreEqual(expected1, actual1);
            ////皆可
            var input2 = "皆可";
            var expected2 = string.Empty;
            var actual2 = Talent.GetInstance().ValidCooperationMode(input2);
            Assert.AreEqual(expected2, actual2);
            ////兼職
            var input3 = "兼職";
            var expected3 = "合作狀態須為\"全職\"or\"合約\"or\"皆可\"";
            var actual3 = Talent.GetInstance().ValidCooperationMode(input3);
            Assert.AreEqual(expected3, actual3);
            ////""
            var input4 = string.Empty;
            var expected4 = "合作狀態須為\"全職\"or\"合約\"or\"皆可\"";
            var actual4 = Talent.GetInstance().ValidCooperationMode(input4);
            Assert.AreEqual(expected4, actual4);
        }

        [TestMethod()]
        public void ValidFilePathTest()
        {
            Assert.Fail();
        }

        [TestMethod()]
        public void ValidInsertMemberTest()
        {
            ////hr@is-land.com.tw
            var input = "hr@is-land.com.tw";
            var expected = "帳號已存在";
            var actual = Talent.GetInstance().ValidInsertMember(input);
            Assert.AreEqual(expected, actual);
            ////hr87@is-land.com.tw
            var input1 = "hr87@is-land.com.tw";
            var expected1 = string.Empty;
            var actual1 = Talent.GetInstance().ValidInsertMember(input1);
            Assert.AreEqual(expected1, actual1);
        }

        [TestMethod()]
        public void ValidMemberStatesTest()
        {
            ////啟用
            var input = "啟用";
            var expected = true;
            var actual = Talent.GetInstance().ValidMemberStates(input);
            Assert.AreEqual(expected, actual);
            ////停用
            var input1 = "停用";
            var expected1 = false;
            var actual1 = Talent.GetInstance().ValidMemberStates(input1);
            Assert.AreEqual(expected1, actual1);
        }

        [TestMethod()]
        public void ValidNewPasswordTest()
        {
            ////newPassword=123,newCheckPassword=123,expected=""
            var expected = string.Empty;
            var actual = Talent.GetInstance().ValidNewPassword("123", "123");
            Assert.AreEqual(expected, actual);
            ////newPassword=1234,newCheckPassword=123,expected="新密碼不一致"
            var expected1 = "新密碼不一致";
            var actual1 = Talent.GetInstance().ValidNewPassword("1234", "123");
            Assert.AreEqual(expected1, actual1);
        }

        [TestMethod()]
        public void ValidProjectExperienceDataTest()
        {
            var input = TestData.GetInstance().TestProjectExperienceData1().DataTableToList<Model.ProjectExperience>(); ;
            var expected = "有空的專案經驗";
            var actual = Talent.GetInstance().ValidProjectExperienceData(input);
            Assert.AreEqual(expected, actual);

            var expected1 = string.Empty;
            var actual1 = Talent.GetInstance().ValidProjectExperienceData(new List<Model.ProjectExperience>());
            Assert.AreEqual(expected1, actual1);
        }

        [TestMethod()]
        public void ValidProjectExperienceIsRepeatTest()
        {
            ////面試編號,公司,專案名稱都沒重複
            string[] data1 = { "2", "屏東科技大學", "智慧水產養殖管理系統", "Windows 10", "MS SQL", "學生", "105/07/01", "105/12/14", "C#", "Asp.Net", "say somthing" };
            string[] data2 = { "1", "亦思科技股份有限公司", "人才資料庫系統", "Windows 10", "MS SQL", "軟體工程師", "106/12/11", "106/12/29", "C#", "Asp.Net", "say somthing" };
            var input = TestData.GetInstance().TestProjectExperienceData(data1, data2);
            var expected = string.Empty;
            var actual = Talent.GetInstance().ValidProjectExperienceIsRepeat(input);
            Assert.AreEqual(expected, actual);
            ///面試編號,公司,專案名稱都重複
            data1 = new string[] { "2", "屏東科技大學", "智慧水產養殖管理系統", "Windows 10", "MS SQL", "學生", "105/07/01", "105/12/14", "C#", "Asp.Net", "say somthing" };
            data2 = new string[] { "2", "屏東科技大學", "智慧水產養殖管理系統", "Windows 10", "MS SQL", "學生", "105/07/01", "105/12/14", "C#", "Asp.Net", "say somthing" };
            var input1 = TestData.GetInstance().TestProjectExperienceData(data1, data2);
            var expected1 = "專案資料重複";
            var actual1 = Talent.GetInstance().ValidProjectExperienceIsRepeat(input1);
            Assert.AreEqual(expected1, actual1);
            ///面試編號重複，公司不同，專案重複
            data1 = new string[] { "2", "亦思科技", "智慧水產養殖管理系統", "Windows 10", "MS SQL", "學生", "105/07/01", "105/12/14", "C#", "Asp.Net", "say somthing" };
            data2 = new string[] { "2", "屏東科技大學", "智慧水產養殖管理系統", "Windows 10", "MS SQL", "學生", "105/07/01", "105/12/14", "C#", "Asp.Net", "say somthing" };
            var input2 = TestData.GetInstance().TestProjectExperienceData(data1, data2);
            var expected2 = string.Empty;
            var actual2 = Talent.GetInstance().ValidProjectExperienceIsRepeat(input2);
            Assert.AreEqual(expected2, actual2);
        }

        [TestMethod()]
        public void ValidStatesTest()
        {
            ////追蹤
            var input = "追蹤";
            var expected = string.Empty;
            var actual = Talent.GetInstance().ValidStates(input);
            Assert.AreEqual(expected, actual);
            ////保留
            var input1 = "保留";
            var expected1 = string.Empty;
            var actual1 = Talent.GetInstance().ValidStates(input1);
            Assert.AreEqual(expected1, actual1);
            ////""
            var input3 = string.Empty;
            var expected3 = string.Empty;
            var actual3 = Talent.GetInstance().ValidStates(input3);
            Assert.AreEqual(expected3, actual3);
            ////面談
            var input2 = "面談";
            var expected2 = "合作狀態須為\"追蹤\"or\"保留\"";
            var actual2 = Talent.GetInstance().ValidStates(input2);
            Assert.AreEqual(expected2, actual2);
        }

        [TestMethod()]
        public void UpdateMemberStatesByAccountTest()
        {
            var input1 = "hrDisabled@is-land.com.tw";
            var expected1 = "啟用成功";
            var actual1 = Talent.GetInstance().UpdateMemberStatesByAccount(input1, "啟用");
            Assert.AreEqual(expected1, actual1);

            var input = "hrDisabled@is-land.com.tw";
            var expected = "停用成功";
            var actual = Talent.GetInstance().UpdateMemberStatesByAccount(input, "停用");
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void DelMemberByAccountTest()
        {
            var input = "hrTest@is-land.com.tw";
            var expected = "刪除成功";
            var actual = Talent.GetInstance().DelMemberByAccount(input);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void ValidIsCompanyMailTest()
        {
            var input = "hr";
            var expected = "帳號須為本公司之帳號";
            var actual = Talent.GetInstance().ValidIsCompanyMail(input);
            Assert.AreEqual(expected, actual);

            var input1 = "hr@gmail.com";
            var expected1 = "帳號須為本公司之帳號";
            var actual1 = Talent.GetInstance().ValidIsCompanyMail(input1);
            Assert.AreEqual(expected1, actual1);

            var input2 = "hr@is-land.com.tw";
            var expected2 = string.Empty;
            var actual2 = Talent.GetInstance().ValidIsCompanyMail(input2);
            Assert.AreEqual(expected2, actual2);
        }

        [TestMethod()]
        public void SelectIdByContactTest()
        {
            var expected = 2;
            var actual = Talent.GetInstance().SelectIdByContact(string.Empty, "C#,JAVA", string.Empty, string.Empty, "2017-12-15", "2018-01-01").Count;
            Assert.AreEqual(expected, actual);

            var expected1 = 2;
            var actual1 = Talent.GetInstance().SelectIdByContact(string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Count;
            Assert.AreEqual(expected1, actual1);

            var expected2 = 1;
            var actual2 = Talent.GetInstance().SelectIdByContact(string.Empty, string.Empty, "合約", string.Empty, string.Empty, string.Empty).Count;
            Assert.AreEqual(expected2, actual2);

            var expected3 = 2;
            var actual3 = Talent.GetInstance().SelectIdByContact(string.Empty, string.Empty, "不限", string.Empty, string.Empty, string.Empty).Count;
            Assert.AreEqual(expected3, actual3);
        }

        [TestMethod()]
        public void ValidIdIsAppearTest()
        {
            var input = "1";
            var expected = true;
            var actual = Talent.GetInstance().ValidIdIsAppear(input);
            Assert.AreEqual(expected, actual);

            var input1 = "AA";
            var expected1 = false;
            var actual1 = Talent.GetInstance().ValidIdIsAppear(input1);
            Assert.AreEqual(expected1, actual1);
        }

        [TestMethod()]
        public void ValidCodeIsRepeatTest()
        {
            var input = TestData.GetInstance().TestCodeData();
            var expected = "60146此代碼已存在\n913736此代碼已存在\n";
            var actual = Talent.GetInstance().ValidCodeIsRepeat(input);
            Assert.AreEqual(expected, actual);

            var input1 = TestData.GetInstance().TestCodeData2();
            var expected1 = string.Empty;
            var actual1 = Talent.GetInstance().ValidCodeIsRepeat(input1);
            Assert.AreEqual(expected1, actual1);
        }

        [TestMethod()]
        public void ValidInterviewIdIsAppearTest()
        {
            var input = "2";
            var expected = true;
            var actual = Talent.GetInstance().ValidInterviewIdIsAppear(input);
            Assert.AreEqual(expected, actual);

            var input1 = "AA";
            var expected1 = false;
            var actual1 = Talent.GetInstance().ValidInterviewIdIsAppear(input1);
            Assert.AreEqual(expected1, actual1);
        }

        [TestMethod()]
        public void ValidInterviewInfoDataTest()
        {
            var expected = string.Empty;
            var actual = Talent.GetInstance().ValidInterviewInfoData("軟體工程師", "賴志維", string.Empty, "2017/8/24", string.Empty,
                                                                     string.Empty, string.Empty, string.Empty,
                                                                     string.Empty);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void SelectIdByKeyWordTest()
        {
            var input = "賴,志";
            var expected = new List<string> { "1" };
            var actual = Talent.GetInstance().SelectIdByKeyWord(input);
            Assert.AreEqual(true, expected.SequenceEqual(actual));

            var input1 = "C#,雲林,教室";
            var expected1 = new List<string> { "1", "2" };
            var actual1 = Talent.GetInstance().SelectIdByKeyWord(input1);
            Assert.AreEqual(true, expected1.SequenceEqual(actual1));

            var input2 = "屏東,";
            var expected2 = new List<string>() { "1" };
            var actual2 = Talent.GetInstance().SelectIdByKeyWord(input2);
            Assert.AreEqual(true, expected2.SequenceEqual(actual2));

            var input3 = ",屏東";
            var expected3 = new List<string>() { "1" };
            var actual3 = Talent.GetInstance().SelectIdByKeyWord(input3);
            Assert.AreEqual(true, expected3.SequenceEqual(actual3));

            var input4 = ",";
            var expected4 = new List<string>();
            var actual4 = Talent.GetInstance().SelectIdByKeyWord(input4);
            Assert.AreEqual(true, expected4.SequenceEqual(actual4));

            var input5 = string.Empty;
            var expected5 = new List<string> { "1", "2" };
            var actual5 = Talent.GetInstance().SelectIdByKeyWord(input5);
            Assert.AreEqual(true, expected5.SequenceEqual(actual5));
        }

        [TestMethod()]
        public void SelectIdByInterviewFilterTest()
        {
            var expected = new List<string>() { "1", "2" };
            var actual = Talent.GetInstance().SelectIdByInterviewFilter("不限", "不限", string.Empty, string.Empty);
            Assert.AreEqual(true, expected.SequenceEqual(actual));

            var expected1 = new List<string>() { "1" };
            var actual1 = Talent.GetInstance().SelectIdByInterviewFilter("已面談", "不限", string.Empty, "2017-08-24");
            Assert.AreEqual(true, expected1.SequenceEqual(actual1));
        }

        [TestMethod()]
        public void SelectContactSituationDataByIdTest()
        {
            var input = "1";
            var expected = 3;
            var actual = Talent.GetInstance().SelectContactSituationDataById(input).Tables.Count;
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void DelTalentByIdTest()
        {
            var input = "3";
            var expected = "刪除成功";
            var actual = Talent.GetInstance().DelTalentById(input);
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void DataClassificationTest()
        {
            var input = TestData.GetInstance().TestCodeData3();
            var expected = 3;
            var actual = Talent.GetInstance().DataDataClassification(input);
            Assert.AreEqual(expected, actual.Tables.Count);

            var input1 = TestData.GetInstance().TestCodeData2();
            var expected1 = 3;
            var actual1 = Talent.GetInstance().DataDataClassification(input1);
            Assert.AreEqual(expected1, actual1.Tables.Count);

            var input2 = TestData.GetInstance().TestCodeData();
            var expected2 = 3;
            var actual2 = Talent.GetInstance().DataDataClassification(input2);
            Assert.AreEqual(expected2, actual2.Tables.Count);
        }

        [TestMethod()]
        public void UpdateContactSituationTest()
        {
            var input = TestData.GetInstance().TestUpdateContactSituationData();
            var expected = "修改成功";
            var actual = Talent.GetInstance().UpdateData(input, "Contact_Situation");
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void DelContactSituationTest()
        {
            var input = TestData.GetInstance().TestDelContactSituationData();
            var expected = "刪除成功";
            var actual = Talent.GetInstance().DelData(input, "Contact_Situation");
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void ContactSituationActionTest()
        {
            var input = TestData.GetInstance().TestContactSituationActionData();
            var expected = "聯繫資料儲存成功";
            var actual = Talent.GetInstance().ContactSituationAction(input, "2");
            Assert.AreEqual(expected, actual);

            var input1 = TestData.GetInstance().TestContactSituationActionData();
            var expected1 = "此聯繫狀況資料，沒有對應的聯繫資料";
            var actual1 = Talent.GetInstance().ContactSituationAction(input, "AA");
            Assert.AreEqual(expected1, actual1);
        }

        [TestMethod()]
        public void CodeActionTest()
        {
            var input = TestData.GetInstance().TestCodeActionData();
            var expected = "代碼儲存成功";
            var actual = Talent.GetInstance().CodeAction(input, "2");
            Assert.AreEqual(expected, actual);

            var input1 = TestData.GetInstance().TestCodeActionData();
            var expected1 = "此代碼，沒有對應的聯繫資料";
            var actual1 = Talent.GetInstance().CodeAction(input, "AA");
            Assert.AreEqual(expected1, actual1);

            var input2 = TestData.GetInstance().TestCodeActionData1();
            var expected2 = "TEST此代碼已存在\n";
            var actual2 = Talent.GetInstance().CodeAction(input, "2");
            Assert.AreEqual(expected2, actual2);
        }

        [TestMethod()]
        public void InsertInterviewInfoDataTest()
        {
            var input = TestData.GetInstance().TestInsertInterviewInfoData();
            var expected = "新增成功";
            var actual = Talent.GetInstance().InsertInterviewInfoData(input, "3");
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void UpdateInterviewInfoDataTest()
        {
            /*var input = TestData.GetInstance().TestUpdateInterviewInfoData();
            var expected = "修改成功";
            var actual = Talent.GetInstance().UpdateInterviewInfoData(input, "4");
            Assert.AreEqual(expected, actual);*/
            Assert.Fail();
        }

        [TestMethod()]
        public void SaveEducationTest()
        {
            ////不存在的欄位
            var input1 = TestData.GetInstance().TestEducationData1();
            var expected1 = "儲存失敗";
            var actual1 = Talent.GetInstance().SaveEducation(input1, "2");
            Assert.AreEqual(expected1, actual1);
            ////沒有對應的面試ID
            var input2 = TestData.GetInstance().TestEducationData1();
            var expected2 = "沒有對應的面試基本資料";
            var actual2 = Talent.GetInstance().SaveEducation(input2, "AA");
            Assert.AreEqual(expected2, actual2);

            var input = TestData.GetInstance().TestEducationData();
            var expected = "儲存成功";
            var actual = Talent.GetInstance().SaveEducation(input, "2");
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void SaveWorkExperienceTest()
        {
            ////不存在的欄位
            var input1 = TestData.GetInstance().TestWorkExperienceData1();
            var expected1 = "儲存失敗";
            var actual1 = Talent.GetInstance().SaveWorkExperience(input1, "2");
            Assert.AreEqual(expected1, actual1);
            ////沒有對應的面試ID
            var input2 = TestData.GetInstance().TestWorkExperienceData();
            var expected2 = "沒有對應的面試基本資料";
            var actual2 = Talent.GetInstance().SaveWorkExperience(input2, "AA");
            Assert.AreEqual(expected2, actual2);

            var input = TestData.GetInstance().TestWorkExperienceData();
            var expected = "儲存成功";
            var actual = Talent.GetInstance().SaveWorkExperience(input, "2");
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void SaveInterviewResultTest()
        {
            var input = TestData.GetInstance().TestSaveInterviewResultData();
            var expected = "儲存成功";
            var actual = Talent.GetInstance().SaveInterviewResult(input, "1");
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void SelectTalentInfoByIdTest()
        {
            var actual = Talent.GetInstance().SelectTalentInfoById(new List<string> { "1", "3" });
        }

        [TestMethod()]
        public void SaveCodeTest()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Code_Id");

            var excepted1 = "此代碼，沒有對應的聯繫資料";
            var actual1 = Talent.GetInstance().SaveCode(dt, "87");
            Assert.AreEqual(excepted1, actual1);

            dt.Clear();
            var excepted2 = "儲存成功";
            var actual2 = Talent.GetInstance().SaveCode(dt, "2");
            Assert.AreEqual(excepted2, actual2);

            dt.Rows.Add("987654321");
            dt.Rows.Add("98765432123456789");
            var excepted = "儲存成功";
            var actual = Talent.GetInstance().SaveCode(dt, "2");
            Assert.AreEqual(excepted, actual);
        }

        [TestMethod()]
        public void SaveContactSituationTest()
        {
            var input = TestData.GetInstance().TestContactSituationActionData();
            var expected = "儲存成功";
            var actual = Talent.GetInstance().SaveContactSituation(input, "2");
            Assert.AreEqual(expected, actual);

            var input1 = TestData.GetInstance().TestContactSituationActionData();
            var expected1 = "此聯繫狀況資料，沒有對應的聯繫資料";
            var actual1 = Talent.GetInstance().SaveContactSituation(input, "AA");
            Assert.AreEqual(expected1, actual1);
        }

        [TestMethod()]
        public void DelInterviewDataTest()
        {
            var expected = "此面談資料不存在";
            var actual = Talent.GetInstance().DelInterviewDataByInterviewId("AA");
            Assert.AreEqual(expected, actual);
        }

        [TestMethod()]
        public void UpLoadImageTest()
        {
            string input1 = @".\images\fdsda.png";
            var expected1 = "不存在的路徑";
            var actual1 = Talent.GetInstance().UpLoadImage(input1, "4");
            Assert.AreEqual(expected1, actual1);

            string input = @".\images\delete.png";
            var expected = @".\images\4.png";
            var actual = Talent.GetInstance().UpLoadImage(input, "4");
            Assert.AreEqual(expected, actual);

            string input2 = @".\images\edit.png";
            var expected2 = @".\images\4.png";
            var actual2 = Talent.GetInstance().UpLoadImage(input2, "4");
            Assert.AreEqual(expected2, actual2);

            string input3 = @".\images\4.png";
            var expected3 = @".\images\4.png";
            var actual3 = Talent.GetInstance().UpLoadImage(input3, "4");
            Assert.AreEqual(expected3, actual3);

            string input4 = string.Empty;
            var expected4 = @"上傳失敗";
            var actual4 = Talent.GetInstance().UpLoadImage(input4, "4");
            Assert.AreEqual(expected4, actual4);
        }

        [TestMethod()]
        public void DelTest()
        {
            var input = @".\images1\4.png";
            var expected = "圖片刪除成功";
            var actual = Talent.GetInstance().DelImageByInterviewId(input);
            Assert.AreEqual(expected, actual);

            var input1 = string.Empty;
            var expected1 = "圖片刪除成功";
            var actual1 = Talent.GetInstance().DelImageByInterviewId(input);
            Assert.AreEqual(expected1, actual);
        }

        [TestMethod()]
        public void UpdateImageTest()
        {
            ////DB跟UI都沒有照片
            var expected = string.Empty;
            var actual = Talent.GetInstance().UpdateImage(string.Empty,string.Empty,"4");
            Assert.AreEqual(expected,actual);

            ////DB沒有照片，UI有照片
            var expected1 = @".\images\4.png";
            var actual1 = Talent.GetInstance().UpdateImage(string.Empty, @".\images\delete.png", "4");
            Assert.AreEqual(expected1, actual1);

            ////DB有照片，UI有照片
            var expected2 = @".\images\4.png";
            var actual2 = Talent.GetInstance().UpdateImage(@".\images\4.png", @".\images\edit.png", "4");
            Assert.AreEqual(expected2, actual2);

            ////DB有照片，UI有照片(都一樣的路徑)
            var expected3 = @".\images\4.png";
            var actual3 = Talent.GetInstance().UpdateImage(@".\images\4.png", @".\images\4.png", "4");
            Assert.AreEqual(expected3, actual3);

            ////DB有照片，UI沒有照片
            var expected4 = string.Empty;
            var actual4 = Talent.GetInstance().UpdateImage(@".\images\4.png", string.Empty, "4");
            Assert.AreEqual(expected4, actual4);
        }
    }
}