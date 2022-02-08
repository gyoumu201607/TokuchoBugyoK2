using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace TokuchoBugyoK2
{
    public partial class CustomDateTimePicker : System.Windows.Forms.DateTimePicker
    {
        private Color _backDisabledColor;
        public CustomDateTimePicker() : base()
        {
            InitializeComponent();
            this.SetStyle(ControlStyles.UserPaint, true);
            _backDisabledColor = Color.FromKnownColor(KnownColor.Control);
        }

        //BackColor　色取得
        [Browsable(true)]
        public override Color BackColor
        {
            get { return base.BackColor; }
            set { base.BackColor = value; }
        }

        //デザイン時色選択用
        [Category("Appearance"),
            Description("The background color of the component when disabled")]
        [Browsable(true)]
        public Color BackDisabledColor
        {
            get { return _backDisabledColor; }
            set { _backDisabledColor = value; }
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            Graphics g = this.CreateGraphics();

            //ドロップダウンボタンの描写位置
            Rectangle dropDownRectangle =
               new Rectangle(ClientRectangle.Width - 17, 0, 17, ClientRectangle.Height);
            //ブラシ設定
            SolidBrush bkgBrush;
            //カレンダー
            ComboBoxState visualState;

            //コントロールが有効のとき色塗りつぶし
            if (this.Enabled)
            {
                bkgBrush = new SolidBrush(this.BackColor);
                visualState = ComboBoxState.Normal;
            }
            else
            {
                //無効のときは_backDisabledColorの色
                bkgBrush = new SolidBrush(this._backDisabledColor);
                visualState = ComboBoxState.Disabled;
            }


            //background塗りつぶし
            //g.FillRectangle(bkgBrush, 0, 0, ClientRectangle.Width, ClientRectangle.Height);
            g.FillRectangle(bkgBrush, ClientRectangle);

            //枠線描写
            Pen pen = new Pen(Brushes.Black);
            pen.Width = 1;
            g.DrawRectangle(pen, 0, 0, ClientRectangle.Width - 1, ClientRectangle.Height - 1);

            //text描写
            g.DrawString(this.Text, this.Font, Brushes.Black, 0, 2);

            //ComboBoxRendererでドロップダウンボタン描写
            ComboBoxRenderer.DrawDropDownButton(g, dropDownRectangle, visualState);

            g.Dispose();
            bkgBrush.Dispose();
            base.OnPaint(pe);
        }
    }
}