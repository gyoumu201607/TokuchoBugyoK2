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
    public partial class Popup_MojiChange : Form
    {
        //不具合No1010（744）カタカナ用インデックス追加
        //public string[] ReturnValue = new string[8];
        public string[] ReturnValue = new string[9];
        GlobalMethod GlobalMethod = new GlobalMethod();

        public Popup_MojiChange()
        {
            InitializeComponent();
        }

        private void Popup_MojiChange_Load(object sender, EventArgs e)
        {
            item_MessageText.Text = "";
            item_MessageText.Visible = false;

            string buf = "";


            // 文字変換の設定
            buf = GlobalMethod.GetCommonValue1("CHOUSA_MOJI_KARA");
            if (buf != null && buf == "1")
            {
                item_Wave.Checked = true;   // ～
            }
            buf = GlobalMethod.GetCommonValue1("CHOUSA_MOJI_X0201");
            if (buf != null && buf == "1")
            {
                 item_Eisuu.Checked = true;  // 英数字
            }
            //不具合No1010（744）カタカナ追加
            buf = GlobalMethod.GetCommonValue1("CHOUSA_MOJI_KANA");
            if (buf != null && buf == "1")
            {
                item_katakana.Checked = true;  // カタカナ
            }

            // 文字変換の対象項目
            buf = GlobalMethod.GetCommonValue1("CHOUSA_MOJI_T_HINMEI");
            if (buf != null && buf == "1")
            {
                item_Hinmei.Checked = true; // 品名
            }
            buf = GlobalMethod.GetCommonValue1("CHOUSA_MOJI_T_KIKAKU");
            if (buf != null && buf == "1")
            {
                item_Kikaku.Checked = true; // 規格
            }
            buf = GlobalMethod.GetCommonValue1("CHOUSA_MOJI_T_HOUKOKUBIKOU");
            if (buf != null && buf == "1")
            {
                item_HoukokuBiko.Checked = true; // 報告備考
            }
            buf = GlobalMethod.GetCommonValue1("CHOUSA_MOJI_T_IRAIBIKOU");
            if (buf != null && buf == "1")
            {
                item_IraiBiko.Checked = true; // 依頼備考
            }
        }

        // 閉じる
        private void item_BtnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // 実行
        private void item_BtnTouroku_Click(object sender, EventArgs e)
        {
            // ▼ReturnValue
            // 0:実行結果 1:実行 それ以外実行無し
            // 1:変換対象 0:全角→半角 1:半角→全角
            // 2:～変換   1:変換 それ以外変換無し
            // 3:英数字   1:変換 それ以外変換無し
            // 4:品名     1:対象 それ以外対象外
            // 5:規格     1:対象 それ以外対象外
            // 6:報告備考 1:対象 それ以外対象外
            // 7:依頼備考 1:対象 それ以外対象外
            //不具合No1010（744）
            // 8:カタカナ変換　1:変換 それ以外変換無し

            // 実行結果に1を追加
            ReturnValue[0] = "1";
            // 変換対象　0:全角→半角 1:半角→全角
            if (item_ZenkakuHankaku.Checked)
            {
                ReturnValue[1] = "0";
            }
            else if (item_HankakuZenkaku.Checked)
            {
                ReturnValue[1] = "1";
            }
            // ～
            if (item_Wave.Checked)
            {
                ReturnValue[2] = "1";
            }
            else
            {
                ReturnValue[2] = "0";
            }
            // 英数字
            if (item_Eisuu.Checked)
            {
                ReturnValue[3] = "1";
            }
            else
            {
                ReturnValue[3] = "0";
            }
            // 品名
            if (item_Hinmei.Checked)
            {
                ReturnValue[4] = "1";
            }
            else
            {
                ReturnValue[4] = "0";
            }
            // 規格
            if (item_Kikaku.Checked)
            {
                ReturnValue[5] = "1";
            }
            else
            {
                ReturnValue[5] = "0";
            }
            // 報告備考
            if (item_HoukokuBiko.Checked)
            {
                ReturnValue[6] = "1";
            }
            else
            {
                ReturnValue[6] = "0";
            }
            // 依頼備考
            if (item_IraiBiko.Checked)
            {
                ReturnValue[7] = "1";
            }
            else
            {
                ReturnValue[7] = "0";
            }
            //不具合No1010（744）
            // カタカナ
            if (item_katakana.Checked)
            {
                ReturnValue[8] = "1";
            }
            else
            {
                ReturnValue[8] = "0";
            }

            this.Close();
        }

    }
}
