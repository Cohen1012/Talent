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
        /// 根據Table組合Parameters(刪除)
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="dr"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        private SqlCommand CombinationDelParameters(SqlCommand cmd, DataRow dr, string tableName)
        {
            ////聯繫狀況
            if (tableName == "Contact_Situation")
            {
                cmd.Parameters.Add("@id", SqlDbType.Int).Value = dr["Contact_status_Id"].ToString();
                return cmd;
            }
            ////代碼
            if (tableName == "Code")
            {
                cmd.Parameters.Add("@id", SqlDbType.VarChar).Value = dr["Code_Id"].ToString();
                return cmd;
            }

            return cmd;
        }

        /// <summary>
        /// 根據Table產生對應的刪除SQL
        /// </summary>
        /// <param name="tableName">要刪除的Table</param>
        /// <returns>刪除SQL</returns>
        private string CombinationDelSQL(string tableName)
        {
            string Del = string.Empty;
            ////聯繫狀況
            if (tableName == "Contact_Situation")
            {
                Del = @"delete from Contact_Situation where Contact_status_Id = @id";
                return Del;
            }
            ////代碼
            if (tableName == "Code")
            {
                Del = @"delete from Code where Code_Id = @id";
                return Del;
            }

            return Del;
        }

        /// <summary>
        /// 根據Table組合Parameters(修改)
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="dr"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        private SqlCommand CombinationUpdateParameters(SqlCommand cmd, DataRow dr, string tableName)
        {
            ////聯繫狀況
            if (tableName == "Contact_Situation")
            {
                cmd.Parameters.Add("@contactStatus", SqlDbType.NVarChar).Value = dr["Contact_Status"].ToString();
                cmd.Parameters.Add("@remarks", SqlDbType.NVarChar).Value = dr["Remarks"].ToString();
                cmd.Parameters.Add("@contactDate", SqlDbType.DateTime).Value = dr["Contact_Date"].ToString();
                cmd.Parameters.Add("@id", SqlDbType.Int).Value = dr["Contact_status_Id"].ToString();
                return cmd;
            }
            ////代碼
            if (tableName == "Code")
            {
                cmd.Parameters.Add("@id", SqlDbType.Int).Value = dr["Id"].ToString();
                cmd.Parameters.Add("@CodeId", SqlDbType.VarChar).Value = dr["Code_Id"].ToString();
                return cmd;
            }

            return cmd;
        }

        /// <summary>
        /// 根據Table產生對應的修改SQL
        /// </summary>
        /// <param name="tableName">要修改的Table</param>
        /// <returns>修改SQL</returns>
        private string CombinationUpdateSQL(string tableName)
        {
            string update = string.Empty;
            ////聯繫狀況
            if (tableName == "Contact_Situation")
            {
                update = @"update Contact_Situation set Contact_Status=@contactStatus,Remarks=@remarks,Contact_Date=@contactDate
                                                       where Contact_status_Id=@id";
                return update;
            }
            ////代碼
            if (tableName == "Code")
            {
                update = @"update Code set Code_Id=@CodeId where Id=@Id";
                return update;
            }

            return update;
        }

        /// <summary>
        /// 根據Table產生對應的新增SQL
        /// </summary>
        /// <param name="tableName"></param>
        /// <returns></returns>
        private string CombinationInsertSQL(string tableName)
        {
            string insert = string.Empty;
            ////聯繫狀況
            if (tableName == "Contact_Situation")
            {
                insert = @"insert into Contact_Situation (Contact_Status,Remarks,Contact_Date,Contact_Id)
                                                     values (@contactStatus,@remarks,@contactDate,@contactId)";
                return insert;
            }
            ////代碼
            if (tableName == "Code")
            {
                insert = @"insert into Code (Code_Id,Contact_Id) values (@codeId,@contactId)";
                return insert;
            }

            return insert;
        }

        /// <summary>
        /// 根據Table組合Parameters(新增)
        /// </summary>
        /// <param name="id"></param>
        /// <param name="cmd"></param>
        /// <param name="dr"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        private SqlCommand CombinationInsertParameters(string id, SqlCommand cmd, DataRow dr, string tableName)
        {
            ////聯繫狀況
            if (tableName == "Contact_Situation")
            {
                cmd.Parameters.Add("@contactStatus", SqlDbType.NVarChar).Value = dr["Contact_Status"].ToString();
                cmd.Parameters.Add("@remarks", SqlDbType.NVarChar).Value = dr["Remarks"].ToString();
                cmd.Parameters.Add("@contactDate", SqlDbType.DateTime).Value = dr["Contact_Date"].ToString();
                cmd.Parameters.Add("@contactId", SqlDbType.Int).Value = id;
                return cmd;
            }
            ////代碼
            if (tableName == "Code")
            {
                cmd.Parameters.Add("@codeId", SqlDbType.NVarChar).Value = dr["Code_Id"].ToString();
                cmd.Parameters.Add("@contactId", SqlDbType.Int).Value = id;
                return cmd;
            }

            return cmd;
        }
        
        /// <summary>
        /// 將查詢結果組合成符合格式的DataTable
        /// </summary>
        /// <param name="ds">查詢決果</param>
        /// <param name="dt">輸出DataTable</param>
        public static void CombinationGrid(DataSet ds, DataTable dt)
        {
            DataRow row = dt.NewRow();
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                row["Contact_Id"] = dr["Contact_Id"].ToString();
                row["Name"] = dr["Name"].ToString();
                row["UpdateTime"] = dr["UpdateTime"].ToString();
            }

            foreach (DataRow dr in ds.Tables[1].Rows)
            {
                row["Code_Id"] += dr[0].ToString() + "\n";
            }

            foreach (DataRow dr in ds.Tables[2].Rows)
            {
                row["Contact_Date"] = dr["Contact_Date"].ToString();
                row["Contact_Status"] = dr["Contact_Status"].ToString();
                row["Remarks"] = dr["Remarks"].ToString();
                break;
            }

            foreach (DataRow dr in ds.Tables[3].Rows)
            {
                row["Interview_Date"] += dr[0].ToString() + "\n";
            }

            dt.Rows.Add(row);

            DataRow row1 = dt.NewRow();

            if (ds.Tables[2].Rows.Count > 1)
            {
                foreach (DataRow dr in ds.Tables[2].Rows)
                {
                    row1["Contact_Date"] = dr["Contact_Date"].ToString();
                    row1["Contact_Status"] = dr["Contact_Status"].ToString();
                    row1["Remarks"] = dr["Remarks"].ToString();
                }

                dt.Rows.Add(row1);
            }
        }
    }
}
