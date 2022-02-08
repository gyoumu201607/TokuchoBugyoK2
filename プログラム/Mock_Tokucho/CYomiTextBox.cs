using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;
//参照に「Microsoft.VisualBasic」を追加

namespace TokuchoBugyoK2
{
    public class YomiTextBox : TextBox
    {
        private const int WM_IME_COMPOSITION = 0x010F;
        private const int GCS_RESULTREADSTR = 0x0200;

        // デリゲートの宣言
        public delegate void CompositionCompletedEventHandler(object sender, CompositionCompletedEventArgs e);

        // イベントハンドラの宣言
        public event CompositionCompletedEventHandler CompositionCompleted;

        [DllImport("Imm32.dll")]
        private static extern int ImmGetContext(IntPtr hWnd);
        [DllImport("Imm32.dll")]
        private static extern int ImmGetCompositionString(int hIMC, int dwIndex, StringBuilder lpBuf, int dwBufLen);
        [DllImport("Imm32.dll")]
        private static extern bool ImmReleaseContext(IntPtr hWnd, int hIMC);

        //-------------------------------------------------------------------------------------
        // コンストラクタ
        public YomiTextBox() : base()
        {
        }

        //-------------------------------------------------------------------------------------
        protected override void WndProc(ref Message m)
        {
            if (CompositionCompleted != null && m.Msg == WM_IME_COMPOSITION)
            {
                if (((int)m.LParam & GCS_RESULTREADSTR) == GCS_RESULTREADSTR)
                {
                    int hImc = ImmGetContext(this.Handle);

                    // 読みの文字数を取得する
                    int size = ImmGetCompositionString(hImc, GCS_RESULTREADSTR, null, 0);

                    StringBuilder sb = new StringBuilder(size);
                    sb.Append(new string(' ', size)); // 受け取る文字列の長さに設定する（これがないと 1 文字だけの時にゴミが入るから）

                    // 読みを取得する
                    ImmGetCompositionString(hImc, GCS_RESULTREADSTR, sb, sb.Length);

                    ImmReleaseContext(this.Handle, hImc);

                    CompositionCompletedEventArgs args = new CompositionCompletedEventArgs(sb.ToString());

                    // CompositionCompleted イベントを発生させる
                    CompositionCompleted(this, args);
                }
            }

            base.WndProc(ref m);
        }
    } // end of YomiTextBox class

    //**************************************************************************************
    // CompositionCompletedEventArgs class（文字の変換を終了したときに発生するイベントのデータ）
    //**************************************************************************************
    [Serializable]
    public class CompositionCompletedEventArgs : EventArgs
    {
        private string FHanKana;  // 半角カタカナ
        private string FKatakana; // 全角カタカナ
        private string FHiragana; // 全角ひらがな

        public string HanKana { get { return FHanKana; } }
        public string Katakana { get { return FKatakana; } }
        public string Hiragana { get { return FHiragana; } }

        public CompositionCompletedEventArgs(string s)
        {
            FHanKana = s;
            FKatakana = Strings.StrConv(s, VbStrConv.Wide, 0); // 半角カナを全角カナに変換する
            FHiragana = Strings.StrConv(FKatakana, VbStrConv.Hiragana, 0); // カタカナをひらがなに変換する
        }

    } // end of CompositionCompletedEventArgs class
}
