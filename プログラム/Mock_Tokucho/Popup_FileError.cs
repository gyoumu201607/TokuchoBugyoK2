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
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using DataTable = System.Data.DataTable;

namespace TokuchoBugyoK2
{
    public partial class Popup_FileError : Form
    {
        GlobalMethod GlobalMethod = new GlobalMethod();
        string ErrorID;
        int Count;
        DataTable ListData = new DataTable();
        public string chousaHinmokuErrorCnt = "";
        public string FileReadErrorTokuchoBangou = "";

        //課題No1300（994）　VIPS
        bool isManyFile=false;
        List<int> FileReadErrorReadCounts = new List<int>();
        public Popup_FileError(string ErrorID, List<int> ReadCounts)
        {
            isManyFile = true;
            this.ErrorID = ErrorID;
            this.FileReadErrorReadCounts = ReadCounts;
            InitializeComponent();
        }

        public Popup_FileError(string ErrorID,int Count)
        {
            this.ErrorID = ErrorID;
            this.Count = Count;
            InitializeComponent();
        }

        private void Popup_FileError_Load(object sender, EventArgs e)
        {

            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            Get_ErrorData();
        }

        private void Get_ErrorData()
        {
            ListData.Clear();
            //課題No1300（994）　VIPS
            if (isManyFile)
            {
                ListData = GlobalMethod.getError(ErrorID, FileReadErrorReadCounts, GlobalMethod.ChangeSqlText(textBox111.Text, 0, 0), chousaHinmokuErrorCnt, FileReadErrorTokuchoBangou);
            }
            else
            {
                ListData = GlobalMethod.getError(ErrorID, Count, GlobalMethod.ChangeSqlText(textBox111.Text, 0, 0), chousaHinmokuErrorCnt, FileReadErrorTokuchoBangou);
            }
            

            c1FlexGrid1.Rows.Count = 1;
            for (int i = 0; i < ListData.Rows.Count; i++)
            {
                c1FlexGrid1.Rows.Add();
                c1FlexGrid1.Rows[i + 1][0] = ListData.Rows[i][0].ToString();
                c1FlexGrid1.Rows[i + 1][1] = ListData.Rows[i][1].ToString();
                c1FlexGrid1.Rows[i + 1][2] = ListData.Rows[i][2].ToString();
                c1FlexGrid1.Rows[i + 1][3] = ListData.Rows[i][3].ToString();
            }

            if (c1FlexGrid1.Rows.Count == 0)
            {
                textBox111.ReadOnly = true;
                pictureBox1.Enabled = false;
            }
        }

        private void textBox111_Enter(object sender, EventArgs e)
        {
            Get_ErrorData();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Application ExcelApp = new Application();
            ExcelApp.DisplayAlerts = false;
            Workbook wb = ExcelApp.Workbooks.Add();
            dynamic xlSheet = null;
            xlSheet = wb.Sheets[1];
            Range rgn = xlSheet.Cells;

            try
            {

                rgn[1, 1] = "エラー発生行";
                rgn[1, 1].Font.Bold = true;
                rgn[1, 2] = "エラー内容";
                rgn[1, 2].Font.Bold = true;
                rgn[1, 3] = "ファイル名";
                rgn[1, 3].Font.Bold = true;
                rgn[1, 4] = "発生日時";
                rgn[1, 4].Font.Bold = true;
                /*
                for (int i = 1; i <= ListData.Rows.Count; i++)
                {
                    rgn[i + 1, 1] = ListData.Rows[i - 1][0].ToString();
                    rgn[i + 1, 2] = ListData.Rows[i - 1][1].ToString();
                }
                */
                String[,] mData2 = new string[ListData.Rows.Count, ListData.Columns.Count];
                for (int i = 0; i < ListData.Rows.Count; i++)
                {
                    for (int k = 0; k < ListData.Columns.Count; k++)
                    {
                        mData2[i, k] = ListData.Rows[i][k].ToString();
                    }
                }

                rgn = xlSheet.Range("A2:D" + (ListData.Rows.Count + 1));
                rgn.Value = mData2;

                string ExcelPath = "Work";
                string ExcelName;
                //課題No1300（994）　VIPS
                if (isManyFile)
                {
                    string bufCount = "";
                    foreach(int num in FileReadErrorReadCounts)
                    {
                        if (bufCount != "")
                        {
                            bufCount += "_";
                        }
                        bufCount += num.ToString();
                    }
                    //ファイル存在しないエラー等だと、番号がつかない（DBに何もいないためゼロで
                    if (bufCount == "")
                    {
                        bufCount = "0";
                    }
                    ExcelName = "FileReadError_" + ErrorID + "_" + bufCount + @".xlsx";
                }
                else
                {
                    ExcelName = "FileReadError_" + ErrorID + "_" + Count + @".xlsx";
                }
                    

                try
                {
                    GlobalMethod.Get_WorkFolder();
                    wb.SaveAs(System.IO.Path.Combine(new System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).Directory.FullName, ExcelPath, ExcelName));



                    Popup_Download form = new Popup_Download();
                    form.TopLevel = false;
                    this.Controls.Add(form);
                    form.ExcelPath = System.IO.Path.Combine(ExcelPath, ExcelName);
                    form.ExcelName = ExcelName;
                    form.Dock = DockStyle.Bottom;
                    form.Show();
                    form.BringToFront();
                }
                catch (Exception)
                {
                    MessageBox.Show("EXCELファイルの作成に失敗いたしました。");
                    return;
                }

                wb.Close(false);
                ExcelApp.Quit();
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                Marshal.ReleaseComObject(xlSheet);
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(ExcelApp);
                xlSheet = null;
                wb = null;
                ExcelApp = null;
                GC.Collect();
            }
        }

        private void textBox111_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void textBox111_TextChanged(object sender, EventArgs e)
        {
            Get_ErrorData();
        }
    }
}
