using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TokuchoBugyoK2
{
    public partial class Popup_MessageBox : Form
    {
        private string _input_text = "";

        public Popup_MessageBox()
        {
            InitializeComponent();
        }

        public Popup_MessageBox(string title, string message, string input_label)
        {
            InitializeComponent();
            if(string.IsNullOrEmpty(title) == false)
                this.Text = title;
            if (string.IsNullOrEmpty(message) == false)
                lblMessage.Text = message;
            if (string.IsNullOrEmpty(input_label) == false)
                lblInput.Text = input_label;
        }


        private void btnOK_Click(object sender, EventArgs e)
        {
            //ダイアログの戻り値をOKに設定
            this.DialogResult = DialogResult.OK;

            //入力内容を取得
            _input_text = this.txtInput.Text;

            //ダイアログを閉じる
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            //ダイアログの戻り値をキャンセルに設定
            this.DialogResult = DialogResult.Cancel;

            //ダイアログを閉じる
            this.Close();
        }

        public string GetInputText()
        {
            return _input_text;
        }
    }
}
