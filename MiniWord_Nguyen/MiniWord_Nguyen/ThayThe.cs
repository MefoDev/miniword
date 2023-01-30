using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MiniWord_Nguyen
{
    public partial class ThayThe : Form
    {
        public ThayThe(RichTextBox rtxghi)
        {
            InitializeComponent();
            rtxSearch = rtxghi;
        }

        private void ThayThe_Load(object sender, EventArgs e)
        {

        }
        private RichTextBox rtxSearch;
        private RichTextBoxFinds GetOptions()
        {
            RichTextBoxFinds rtbf = new RichTextBoxFinds();
            rtbf = rtbf | RichTextBoxFinds.MatchCase;
            return rtbf;
        }
        private int pos = 0;
        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                string search = txtSearch.Text;

                pos = rtxSearch.Find(search, pos, GetOptions());
                if (pos == -1)
                {
                    rtxSearch.Select(0, 0);
                    pos = 0;
                    MessageBox.Show("Search is Completed");
                }
                rtxSearch.Focus();
                pos += 1;
            }
            catch
            {

            }

        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string Replace = txtReplace.Text;
                if(Replace == "")
                {
                    Replace = " ";
                }
                if(rtxSearch.SelectedText.Length != 0)
                {
                    rtxSearch.SelectedText = rtxSearch.SelectedText.Replace(rtxSearch.SelectedText, Replace);
                }    
                BtnSearch.PerformClick(); // chọn vùng bôi đen để thay thế
            }
            catch
            {

            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if(txtSearch.Text.Length != 0) rtxSearch.Text = rtxSearch.Text.Replace(txtSearch.Text, txtReplace.Text);

        }
    }
}
