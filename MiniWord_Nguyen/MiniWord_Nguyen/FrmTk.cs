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
    public partial class frmTK : Form
    {
        private RichTextBox rtxsearch;
        public frmTK(RichTextBox rtxghi) 
        {
            InitializeComponent();
            rtxsearch = rtxghi;
        }

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
                
                pos = rtxsearch.Find(search, pos, GetOptions());
                if(pos == -1)
                {
                    rtxsearch.Select(0, 0);
                    pos = 0;
                    MessageBox.Show("Search is Completed");
                }
                rtxsearch.Focus();
                pos += 1;
            }
            catch
            {

            }


        }
    }
}
