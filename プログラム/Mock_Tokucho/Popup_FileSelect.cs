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
    public partial class Popup_FileSelect : Form
    {
        GlobalMethod GlobalMethod = new GlobalMethod();
        string ErrorID;
        int Count;
        DataTable ListData = new DataTable();
        public string chousaHinmokuErrorCnt = "";
        public string FileReadErrorTokuchoBangou = "";

        public Popup_FileSelect(string ErrorID,int Count)
        {
            this.ErrorID = ErrorID;
            this.Count = Count;
            InitializeComponent();
        }

        private void Popup_FileSelect_Load(object sender, EventArgs e)
        {

            

            
        }

        
       
        
    }
}
