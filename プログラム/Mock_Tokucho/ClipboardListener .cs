using System.Windows.Forms;
using System.Runtime.InteropServices;
using System;

namespace TokuchoBugyoK2
{
    public class ClipboardListener : NativeWindow
    {
        [DllImport("user32")]
        public static extern IntPtr SetClipboardViewer(IntPtr hWndNewViewer);

        [DllImport("user32")]
        public static extern bool ChangeClipboardChain(IntPtr hWndRemove, IntPtr hWndNewNext);

        [DllImport("user32")]
        public extern static int SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        private const int WM_DRAWCLIPBOARD = 0x0308;
        private const int WM_CHANGECBCHAIN = 0x030D;

        private IntPtr nextHandle = IntPtr.Zero;
        public event EventHandler DrawClipboard;
        private Form handlerF = null;
        private int winCount = 0;

        public ClipboardListener(Form f)
        {
            f.HandleCreated += OnHandleCreated;
            f.HandleDestroyed += OnHandleDestroyed;
            handlerF = f;
        }

        internal void OnHandleCreated(object sender, EventArgs e)
        {
            //if (handlerF.Name.Equals("Madoguchi_Input"))
            //{
            //    if (((Madoguchi_Input)handlerF).copyData != null)
            //        ((Madoguchi_Input)handlerF).copyData.Clear();
            //}
            // NativeWindowクラスへのForm登録(メッセージフック開始)
            AssignHandle(((Form)sender).Handle);

            // クリップボードチェインに登録
            nextHandle = SetClipboardViewer(this.Handle);
        }

        internal void OnHandleDestroyed(object sender, EventArgs e)
        {
            //if (handlerF.Name.Equals("Madoguchi_Input"))
            //{
            //    if (((Madoguchi_Input)handlerF).copyData != null)
            //        ((Madoguchi_Input)handlerF).copyData.Clear();
            //}
            // クリップボードチェインから削除
            ChangeClipboardChain(this.Handle, nextHandle);

            // NativeWindowクラスの後始末(Formに対してのメッセージフック解除)
            ReleaseHandle();
        }

        protected override void WndProc(ref Message msg)
        {
            if (handlerF.Name.Equals("Madoguchi_Input"))
            {
                if (winCount > 0)
                {
                    winCount = 0;
                    return;
                }
                if (((Madoguchi_Input)handlerF).isFormCopy)
                    winCount++;
                else
                    winCount = 0;
            }
            if (handlerF.Name.Equals("Tokumei_Input"))
            {
                if (winCount > 0)
                {
                    winCount = 0;
                    return;
                }
                if (((Tokumei_Input)handlerF).isFormCopy)
                    winCount++;
                else
                    winCount = 0;
            }
            if (handlerF.Name.Equals("Jibun_Input"))
            {
                if (winCount > 0)
                {
                    winCount = 0;
                    return;
                }
                if (((Jibun_Input)handlerF).isFormCopy)
                    winCount++;
                else
                    winCount = 0;
            }
            if (msg.Msg == WM_DRAWCLIPBOARD)
            {
                if (nextHandle != IntPtr.Zero)
                {
                    SendMessage(nextHandle, msg.Msg, msg.WParam, msg.LParam);
                }
                DrawClipboard?.Invoke(this, new EventArgs());
            }
            if (msg.Msg == WM_CHANGECBCHAIN)
            {
                if (msg.WParam == nextHandle)
                {
                    // WParamが次のウィンドウのハンドルと同じなら
                    // 次のウィンドウはクリップボードビューアチェインから外れたことになる。
                    // 今後はLPARAM のハンドルに対してメッセージを送るため nextHandleを変更する
                    nextHandle = msg.LParam;
                }
                else if (nextHandle != IntPtr.Zero)
                {
                    // WPARAM が次のウィンドウでなければ、
                    // そのまま次のウィンドウにWM_CHANGECBCHAIN メッセージを送る
                    SendMessage(nextHandle, msg.Msg, msg.WParam, msg.LParam);
                }
            }
            base.WndProc(ref msg);
        }
    }
}