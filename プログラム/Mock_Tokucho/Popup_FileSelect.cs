using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using DataTable = System.Data.DataTable;

namespace TokuchoBugyoK2
{
    public partial class Popup_FileSelect : Form
    {
        //デバイス設定のグリッド列番号
        private enum GRID_COL : int
        {
            AVAILABLE = 0
            , FILE_NAME
            
        }
        GlobalMethod GlobalMethod = new GlobalMethod();
        private string _FolderPath;
        private List<string> _selectFiles = new List<string>();

        public List<string> selectFiles
        {
            get
            {
                return _selectFiles;
            }
        }
        
        public Popup_FileSelect(string FolderPath)
        {
            _FolderPath = FolderPath;
            InitializeComponent();
        }

        private void Popup_FileSelect_Load(object sender, EventArgs e)
        {
            //Grid初期化
            gridFileList.Rows.Count = 1;

            if (Directory.Exists(_FolderPath)==true)
            {
                //MessageBox.Show("フォルダ存在");
                //グリッドにファイル名一覧セット
                //指定ディレクトリ
                DirectoryInfo di = new System.IO.DirectoryInfo(_FolderPath);

                FileInfo[] files =
                    di.GetFiles("*.xlsm", System.IO.SearchOption.AllDirectories);
                foreach(FileInfo fi in files)
                {
                    //隠しファイルはGridのセットしない
                    if((fi.Attributes & FileAttributes.Hidden)== FileAttributes.Hidden)
                    {
                        //Debug用
                        Console.WriteLine(fi.Name);
                    }
                    else
                    {
                        gridFileList.Rows.Count++;
                        gridFileList.Rows[gridFileList.Rows.Count - 1][(int)GRID_COL.AVAILABLE] = false;
                        gridFileList.Rows[gridFileList.Rows.Count - 1][(int)GRID_COL.FILE_NAME] = fi.Name;
                    }
                    
                }
                this.checkBox1.Checked = true;
            }

            
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            //念のため空リスト作成しておく
            _selectFiles = new List<string>();
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            //チェックされているファイルパスをListに格納
            _selectFiles = new List<string>();
            for (int i = 1; i < gridFileList.Rows.Count; i++)
            {
                if ((bool)gridFileList.Rows[i][(int)GRID_COL.AVAILABLE] == true)
                {
                    _selectFiles.Add(gridFileList.Rows[i][(int)GRID_COL.FILE_NAME].ToString());
                }
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            bool bolcheck;
            if (checkBox1.Checked)
            {
                bolcheck = true;
            }
            else
            {
                bolcheck = false;
            }
            for(int i=1; i < gridFileList.Rows.Count; i++)
            {
                gridFileList.Rows[i][(int)GRID_COL.AVAILABLE] = bolcheck;
            }
        }
    }
}
