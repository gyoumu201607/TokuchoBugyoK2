using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TokuchoBugyoK2
{
    public partial class CustomDateTimePicker1 : System.Windows.Forms.DateTimePicker
    {
        public CustomDateTimePicker1()
        {
            InitializeComponent();
        }


        protected override void OnPaint(PaintEventArgs pe)
        {
            base.OnPaint(pe);
        }

        const int WM_ERASEBKGND = 0x14;
        const int WM_ERASEBKGND2 = 0x20;
        const int WM_PAINT = 0x000F;
        protected override void WndProc(ref Message m)
        {
            /*
            Console.WriteLine("★★TEST");
            base.WndProc(ref m);
            if (m.Msg == WM_ERASEBKGND)
            {
                Console.WriteLine("★★OK");
                Redraw();
            }*/
            if (m.Msg == WM_ERASEBKGND || m.Msg == WM_ERASEBKGND2)
            {
                try
                {
                    using (var g = Graphics.FromHdc(m.WParam))
                    {
                        using (var b = new SolidBrush(this.BackColor))
                        {
                            g.FillRectangle(b, ClientRectangle);
                        }
                    }
                    Console.WriteLine(m.Msg.ToString());
                    return;
                }
                catch (Exception)
                {
                    Console.WriteLine(m.Msg.ToString());
                }
            }
            base.WndProc(ref m);

        }
        public override Color BackColor
        {
            get { return base.BackColor; }
            set { base.BackColor = value; }
        }

        public Bitmap GetControlImage()
        {
            var bmp = new Bitmap(this.Width, this.Height);
            this.DrawToBitmap(bmp, new Rectangle(0, 0, this.Width, this.Height));

            using (var g = Graphics.FromImage(bmp))
            {
                ColorMap[] cm = { new ColorMap() };
                cm[0].OldColor = SystemColors.Window;
                cm[0].NewColor = this.BackColor;

                ImageAttributes ia = new ImageAttributes();
                ia.SetRemapTable(cm);

                g.DrawImage(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height), 0, 0, bmp.Width, bmp.Height, GraphicsUnit.Pixel, ia);
            }
            return bmp;
        }

        private void Redraw()
        {
            Size bsz = SystemInformation.Border3DSize;

            using (var g = CreateGraphics())
            {
                using (Bitmap bmp = GetControlImage())
                {
                    g.DrawImage(bmp, -bsz.Width + 2, -bsz.Height + 2);
                }
            }
        }
    }
}
