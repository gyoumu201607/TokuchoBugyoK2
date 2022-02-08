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
    public partial class Popup_KoujiJimusyo : Form
    {
        public Popup_KoujiJimusyo()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            c1FlexGrid1.Rows.Add();
            c1FlexGrid1.Rows[c1FlexGrid1.Rows.Count - 1][3] = textBox7.Text;
            c1FlexGrid1.Rows[c1FlexGrid1.Rows.Count - 1][4] = textBox10.Text;
            c1FlexGrid1.Rows[c1FlexGrid1.Rows.Count - 1][5] = textBox9.Text;
            c1FlexGrid1.Rows[c1FlexGrid1.Rows.Count - 1][6] = textBox8.Text;
            c1FlexGrid1.Rows[c1FlexGrid1.Rows.Count - 1][7] = textBox1.Text;
            c1FlexGrid1.Rows[c1FlexGrid1.Rows.Count - 1][8] = textBox3.Text;
            c1FlexGrid1.Rows[c1FlexGrid1.Rows.Count - 1][9] = textBox4.Text;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int _row = 0;
            for (int i = 0; i < c1FlexGrid1.Rows.Count - 1; i++)
            {
                if (c1FlexGrid1.Rows[c1FlexGrid1.Rows.Count - 1][3].ToString() == textBox7.Text)
                {
                    _row = i;
                    break;
                }
            }

            if (_row == 0)
            {

            }
            else
            {
                c1FlexGrid1.Rows[_row][3] = textBox7.Text;
                c1FlexGrid1.Rows[_row][4] = textBox10.Text;
                c1FlexGrid1.Rows[_row][5] = textBox9.Text;
                c1FlexGrid1.Rows[_row][6] = textBox8.Text;
                c1FlexGrid1.Rows[_row][7] = textBox1.Text;
                c1FlexGrid1.Rows[_row][8] = textBox3.Text;
                c1FlexGrid1.Rows[_row][9] = textBox4.Text;
            }
        }

        private void c1FlexGrid1_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {

            var hti = this.c1FlexGrid1.HitTest(new System.Drawing.Point(e.X, e.Y));

            //if (hti.Column == 1 & hti.Row != 0)
            if (hti.Column == 1 & hti.Row > 0)
            {
                var _row = hti.Row;

                textBox7.Text  = c1FlexGrid1.Rows[_row][3].ToString();
                textBox10.Text = c1FlexGrid1.Rows[_row][4].ToString();
                textBox9.Text  = c1FlexGrid1.Rows[_row][5].ToString();
                textBox8.Text  = c1FlexGrid1.Rows[_row][6].ToString();
                textBox1.Text  = c1FlexGrid1.Rows[_row][7].ToString();
                textBox3.Text  = c1FlexGrid1.Rows[_row][8].ToString();
                textBox4.Text  = c1FlexGrid1.Rows[_row][9].ToString();
            }
        }

        private void Popup_KoujiJimusyo_Load(object sender, EventArgs e)
        {
            c1FlexGrid1.Rows[1][3] = "ﾃｽﾄ事務所";
            c1FlexGrid1.Rows[1][4] = "〇〇部署";
            c1FlexGrid1.Rows[1][5] = "○○課長";
            c1FlexGrid1.Rows[1][6] = "担当者さん";
            c1FlexGrid1.Rows[1][7] = "00-0000-0000";
            c1FlexGrid1.Rows[1][8] = "00-0000-0000";
            c1FlexGrid1.Rows[1][9] = "test@test.co.jp";
        }
    }
}
