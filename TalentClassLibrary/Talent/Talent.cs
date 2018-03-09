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
        private static Talent talent = new Talent();

        public static Talent GetInstance() => talent;

        /// <summary>
        /// 紀錄大略發生的錯誤
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 將資料依狀態分類成修改，新增，刪除的DataTable
        /// </summary>
        /// <param name="DataList">資料</param>
        /// <returns>由新增，修改，刪除DataTable組成的DataSet</returns>
        public DataSet DataDataClassification(System.Data.DataTable DataList)
        {
            DataSet ds = new DataSet();
            DataTable insertList = new DataTable();
            DataTable updateList = new DataTable();
            DataTable delList = new DataTable();

            try
            {
                ////需要新增的資料
                var filter = (from row in DataList.AsEnumerable()
                              where row.Field<string>("Flag") == "I"
                              select row);
                if (filter.Any())
                {
                    insertList = filter.CopyToDataTable();
                }
                ////需要修改的資料
                filter = (from row in DataList.AsEnumerable()
                          where row.Field<string>("Flag") == "U"
                          select row);
                if (filter.Any())
                {
                    updateList = filter.CopyToDataTable();
                }
                ////需要刪除的資料
                filter = (from row in DataList.AsEnumerable()
                          where row.Field<string>("Flag") == "D"
                          select row);
                if (filter.Any())
                {
                    delList = filter.CopyToDataTable();
                }
                ds.Tables.Add(insertList);
                ds.Tables.Add(updateList);
                ds.Tables.Add(delList);
                return ds;
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                return new DataSet();
            }
        }

        /// <summary>
        /// 儲存聯繫狀況資料
        /// </summary>
        /// <param name="data">聯繫狀況資料</param>
        /// <param name="id">對應的聯繫ID</param>
        /// <returns>聯繫資料儲存成功or錯誤訊息</returns>
        public string ContactSituationAction(DataTable data, string id)
        {
            DataSet ds = this.DataDataClassification(data);
            string sqlStr = string.Empty;
            if (!this.ValidIdIsAppear(id))
            {
                return "此聯繫狀況資料，沒有對應的聯繫資料";
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
                if (this.InsertData(ds.Tables[0], id, "Contact_Situation") != "新增成功")
                {
                    return "聯繫狀況資料新增失敗";
                }
            }

            if (ds.Tables[1].Rows.Count > 0)
            {
                if (this.UpdateData(ds.Tables[1], "Contact_Situation") != "修改成功")
                {
                    return "聯繫狀況資料修改失敗";
                }
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
                if (this.DelData(ds.Tables[2], "Contact_Situation") != "刪除成功")
                {
                    return "聯繫狀況資料刪除失敗";
                }
            }

            return "聯繫資料儲存成功";
        }

        /// <summary>
        /// 儲存代碼
        /// </summary>
        /// <param name="data">代碼</param>
        /// <param name="id">對應的聯繫ID</param>
        /// <returns>代碼儲存成功or錯誤訊息</returns>
        public string CodeAction(DataTable data, string id)
        {
            DataSet ds = this.DataDataClassification(data);
            string sqlStr = string.Empty;
            if (!this.ValidIdIsAppear(id))
            {
                return "此代碼，沒有對應的聯繫資料";
            }

            string validMsg = this.ValidCodeIsRepeat(data);
            if (!validMsg.Equals(string.Empty))
            {
                return validMsg;
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
                if (this.InsertData(ds.Tables[0], id, "Code") != "新增成功")
                {
                    return "代碼新增失敗";
                }
            }

            if (ds.Tables[1].Rows.Count > 0)
            {
                if (this.UpdateData(ds.Tables[1], "Code") != "修改成功")
                {
                    return "代碼資料修改失敗";
                }
            }

            if (ds.Tables[0].Rows.Count > 0)
            {
                if (this.DelData(ds.Tables[2], "Code") != "刪除成功")
                {
                    return "代碼刪除失敗";
                }
            }

            return "代碼儲存成功";
        }

        /// <summary>
        /// 登入功能
        /// </summary>
        /// <param name="account">帳號</param>
        /// <param name="password">密碼</param>
        /// <returns>登入成功回傳帳號，登入失敗回傳"登入失敗"，該帳號停用中回傳"該帳號停用中"</returns>
        public string SignIn(string account, string password)
        {
            ErrorMessage = string.Empty;
            string msg = "登入失敗";
            account = TalentCommon.GetInstance().AddMailFormat(account);
            password = Common.GetInstance().PasswordEncryption(password.ToLower());
            string select = @"select Account,States from Member where Account=@account and Password = @password";
            DataTable dt = new DataTable();
            try
            {
                using (SqlDataAdapter da = new SqlDataAdapter(select, ScConnection))
                {
                    da.SelectCommand.Parameters.Add("@account", SqlDbType.NVarChar).Value = account;
                    da.SelectCommand.Parameters.Add("@password", SqlDbType.NVarChar).Value = password;

                    da.Fill(dt);

                    if (dt.Rows.Count == 0)
                    {
                        return msg;
                    }
                }
                if (TalentValid.GetInstance().ValidMemberStates(dt.Rows[0][1].ToString()))
                {
                    return dt.Rows[0][0].ToString();
                }
                else
                {
                    return "該帳號停用中";
                }
            }
            catch (Exception ex)
            {
                LogInfo.WriteErrorInfo(ex);
                ErrorMessage = "資料庫發生錯誤";
                return msg;
            }
            finally
            {
                this.CloseDatabaseConnection();
            }
        }

        ///// <summary>
        ///// 將控制項的寬，高，左邊距，頂邊距和字體大小暫存到tag屬性中
        ///// </summary>
        ///// <param name="cons">遞歸控制項中的控制項</param>
        //public void SetTag(Control cons)
        //{
        //    foreach (Control con in cons.Controls)
        //    {
        //        con.Tag = con.Width + ":" + con.Height + ":" + con.Left + ":" + con.Top + ":" + con.Font.Size;
        //        if (con.Controls.Count > 0)
        //            SetTag(con);
        //    }
        //}

        ////根據窗體大小調整控制項大小
        //public void SetControls(float newx, float newy, Control cons)
        //{
        //    //遍歷窗體中的控制項，重新設置控制項的值
        //    foreach (Control con in cons.Controls)
        //    {
        //        string[] mytag = con.Tag.ToString().Split(new char[] { ':' });//獲取控制項的Tag屬性值，並分割後存儲字元串數組
        //        float a = Convert.ToSingle(mytag[0]) * newx;//根據窗體縮放比例確定控制項的值，寬度
        //        con.Width = (int)a;//寬度
        //        a = Convert.ToSingle(mytag[1]) * newy;//高度
        //        con.Height = (int)(a);
        //        a = Convert.ToSingle(mytag[2]) * newx;//左邊距離
        //        con.Left = (int)(a);
        //        a = Convert.ToSingle(mytag[3]) * newy;//上邊緣距離
        //        con.Top = (int)(a);
        //        Single currentSize = Convert.ToSingle(mytag[4]) * newy;//字體大小
        //        con.Font = new Font(con.Font.Name, currentSize, con.Font.Style, con.Font.Unit);
        //        if (con.Controls.Count > 0)
        //        {
        //            SetControls(newx, newy, con);
        //        }
        //    }
        //}
    }
}
