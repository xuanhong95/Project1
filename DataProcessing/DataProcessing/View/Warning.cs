using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataProcessing.View
{
    public partial class Warning : Form
    {
        public Warning()
        {
            InitializeComponent();
            label2.Visible = false;
            textBox1.Visible = false;
        }

        static Warning MsgBox;
        static DialogResult result = DialogResult.No;
        public static int yourchoise = 0;
        public static string newnamecolor = "";

        public static DialogResult Show(string text)
        {
            MsgBox = new Warning();
            MsgBox.label1.Text = text;
            MsgBox.ShowDialog();
            return result;         
        }
        /// <summary>
        /// Button Xóa
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            yourchoise = 1;
            label2.Visible = false;
            textBox1.Visible = false;
        }

        /// <summary>
        /// Button gộp
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            yourchoise = 2;
            label2.Visible = false;
            textBox1.Visible = false;
        }

        /// <summary>
        /// Button đổi tên
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            yourchoise = 3;
            label2.Visible = true;
            textBox1.Visible = true;
        }
        /// <summary>
        /// Tên mã màu mới
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        /// <summary>
        /// Button OK
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (yourchoise == 1)
            {
                this.Visible = false;
            }
            else if (yourchoise == 2)
            {
                this.Visible = false;
            }
            else if(yourchoise == 3)
            {
                newnamecolor = textBox1.Text;
                if (newnamecolor == "")
                {
                    MessageBox.Show("Bạn chưa nhập tên màu mới");
                }
                else
                {
                    this.Visible = false;
                }
            }
            
        }

        /// <summary>
        /// Button Cancel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            yourchoise = 4;
            this.Visible = false;
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        
    }
}
