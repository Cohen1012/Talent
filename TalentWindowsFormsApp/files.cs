using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TalentClassLibrary;
using TalentClassLibrary.Model;
using ShareClassLibrary;
using System.IO;
using System.Reflection;
using Newtonsoft.Json;

namespace TalentWindowsFormsApp
{
    public partial class Files : Form
    {
        /// <summary>
        /// UI附加檔案
        /// </summary>
        private List<FileControl> FileUI = new List<FileControl>();

        /// <summary>
        /// DB附加檔案
        /// </summary>
        private DataTable FileDB = new DataTable();

        /// <summary>
        /// 面談ID
        /// </summary>
        private string InterviewId = string.Empty;

        /// <summary>
        /// 附加檔案模式
        /// </summary>
        private string AttachedFileMode = string.Empty;

        public Files()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 編修模式
        /// </summary>
        /// <param name="interviewId"></param>
        /// <param name="filesMode"></param>
        public Files(string interviewId, string filesMode)
        {
            InitializeComponent();
            InterviewId = interviewId;
            AttachedFileMode = filesMode;
            this.ShowData();
        }

        private void ShowData()
        {
            FileDB = Talent.GetInstance().SelectFiles(InterviewId, AttachedFileMode).Copy();
            if (!string.IsNullOrEmpty(Talent.GetInstance().ErrorMessage))
            {
                MessageBox.Show(Talent.GetInstance().ErrorMessage, "錯誤訊息");
                return;
            }

            List<AttachedFiles> list = FileDB.DataTableToList<AttachedFiles>();
            foreach (AttachedFiles attachedFiles in list)
            {
                FileControl fileControl = new FileControl();
                fileControl.BtnClick += FileControl_BtnClick;
                fileControl.FileFilterClick += FileControl_FileFilterClick;
                fileControl.SetFileInfo(attachedFiles.File_Path, Path.GetFileName(attachedFiles.File_Path));
                FileUI.Add(fileControl);
                splitContainer1.Panel1.Controls.Add(fileControl);

                this.RepositionFiles();

            }
        }

        /// <summary>
        /// 新增一個檔案
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NewFileBtn_Click(object sender, EventArgs e)
        {
            FileControl fileControl = new FileControl();
            fileControl.BtnClick += FileControl_BtnClick;
            fileControl.FileFilterClick += FileControl_FileFilterClick;
            FileUI.Add(fileControl);
            splitContainer1.Panel1.Controls.Add(fileControl);

            this.RepositionFiles();
        }

        /// <summary>
        /// 選取檔案
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FileControl_FileFilterClick(object sender, EventArgs e)
        {
            FileControl fileControl = sender as FileControl;
            OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "請選擇上傳的檔案",
                Filter = @"All Image And PDF Files|*.bmp;*.ico;*.gif;*.jpeg;*.jpg;*.png;*.tif;*.tiff;*.pdf",
                Multiselect = false,
                RestoreDirectory = true
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                foreach (FileControl fc in FileUI)
                {
                    if (ofd.SafeFileName.Equals(fc.GetFileName()))
                    {
                        MessageBox.Show("檔案已存在，請更改檔名，或將覆蓋舊檔案後可上傳", "錯誤訊息");
                        return;
                    }
                }

                fileControl.SetFileInfo(ofd.FileName, ofd.SafeFileName);
            }
        }

        /// <summary>
        /// 定位檔案
        /// </summary>
        private void RepositionFiles()
        {
            for (int i = 0; i < FileUI.Count; i++)
            {
                FileUI[i].Location = new Point(13, 13 + (i * 1) + (i * FileUI[i].Height));
            }
        }

        /// <summary>
        /// 移除檔案
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FileControl_BtnClick(object sender, EventArgs e)
        {
            FileControl fileControl = (FileControl)sender;
            FileUI.Remove(fileControl);
            splitContainer1.Panel1.Controls.Remove(fileControl);
            this.RepositionFiles();
        }

        /// <summary>
        /// 關閉視窗
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CloseBtn_Click(object sender, EventArgs e)
        {
            List<string> fileUIlist = new List<string>();
            foreach (FileControl fileControl in FileUI)
            {
                fileUIlist.Add(fileControl.GetFilePath());
            }
            List<string> fileDBlist = new List<string>();
            foreach (DataRow dr in FileDB.Rows)
            {
                fileDBlist.Add(dr["File_Path"].ToString());
            }

            if (!fileDBlist.SequenceEqual(fileUIlist))
            {
                DialogResult result = MessageBox.Show("附加檔案尚未儲存，您確定要離開嗎?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.No)
                {
                    return;
                }
            }

            this.Close();
        }

        /// <summary>
        /// 儲存附加檔案
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveFilesBtn_Click(object sender, EventArgs e)
        {
            List<string> FilePathList = new List<string>();
            DialogResult result = MessageBox.Show("資料是否正確?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.No)
            {
                return;
            }

            foreach (FileControl fileControl in FileUI)
            {
                if (string.IsNullOrEmpty(fileControl.GetFilePath()))
                {
                    MessageBox.Show("有空白的路徑", "錯誤訊息");
                    return;
                }

                FilePathList.Add(fileControl.GetFilePath());
            }

            string msg = Talent.GetInstance().SaveAttachedFilesToDB(FilePathList, InterviewId, AttachedFileMode);
            MessageBox.Show(msg, "訊息");
            if (msg.Equals("檔案上傳成功"))
            {
                this.DialogResult = DialogResult.OK;
                this.Hide();
                return;
            }
        }
    }
}
