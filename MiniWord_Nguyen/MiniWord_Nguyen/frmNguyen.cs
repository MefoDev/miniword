using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MiniWord_Nguyen
{
    public partial class frmNguyen : Form
    {
        public static bool checksave = false;
        public frmNguyen()
        {
            InitializeComponent();
            FontText();
            tscmbFont.SelectedItem = "Times New Roman";
            tscFontsize.SelectedItem = "11";
            chenEmoji();
        }

        private void FontText()
        {
            try
            {
                foreach (FontFamily f in FontFamily.Families)
                {
                    tscmbFont.Items.Add(f.Name);
                }
            }
            catch
            {
                MessageBox.Show("loi font");
            }
        }
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (checksave == true)
            {
                DialogResult savePrompt = MessageBox.Show("Do you want to save changes to Document?", "SAVE", MessageBoxButtons.YesNoCancel);
                switch (savePrompt)
                {
                    case DialogResult.Cancel:
                        break;
                    case DialogResult.No:
                        open();
                        break;
                    case DialogResult.Yes:
                        saveAsToolStripMenuItem.PerformClick();
                        open();
                        break;
                }
            }
            else
            {
                open();
            }
        }
        
        private void tenfile(string path)
        {
            FileHienTai = true;
            LayFileHienTai = path;
            string[] s = path.Split('\\');
            lbtenfile.Text = s[s.Length - 1];
        }
        private void open()
        {
            try
            {
                panel7.Visible = true;
                rtxGhi.Visible = true;
                OpenFileDialog s = new OpenFileDialog();
                s.Filter = "Open File (*.rtf)|*.rtf|(*.txt)|*.txt";
                if (s.ShowDialog() == DialogResult.OK)
                {
                    string path = s.FileName;
                    if (path != "")
                    {
                        rtxGhi.LoadFile(path);
                        LayFileHienTai = path;
                        tenfile(path);
                        FileHienTai = true;
                    }
                }
            }
            catch
            {
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Clipboard.GetDataObject().GetDataPresent(DataFormats.Text)) rtxGhi.Paste();
        }

        private void zoomInToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (size_PhongTo >= 64) return;
            this.rtxGhi.ZoomFactor = size_PhongTo;
            size_PhongTo += 2;
        }

        private void zoomOutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (size_PhongTo - 2 < 1) return;
            else size_PhongTo -= 2;
            this.rtxGhi.ZoomFactor = size_PhongTo;
        }
        private float size_PhongTo = 1;
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (rtxGhi.SelectionLength > 0) rtxGhi.Cut();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            if (rtxGhi.SelectionLength > 0) rtxGhi.Copy();
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
           rtxGhi.Undo();
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            rtxGhi.Redo();
        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (rtxGhi.SelectionLength > 0) rtxGhi.Cut();
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (rtxGhi.SelectionLength > 0) rtxGhi.Copy();
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Clipboard.GetDataObject().GetDataPresent(DataFormats.Text)) rtxGhi.Paste();
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            rtxGhi.SelectionAlignment = HorizontalAlignment.Left;
            tsblef.BackColor = SystemColors.GradientActiveCaption;
            tsbCenter.BackColor = SystemColors.Control;
            tsbright.BackColor = SystemColors.Control;
        }

        private void toolStripButton12_Click(object sender, EventArgs e)
        {
            rtxGhi.SelectionAlignment = HorizontalAlignment.Center;
            tsblef.BackColor = SystemColors.Control;
            tsbCenter.BackColor = SystemColors.GradientActiveCaption;
            tsbright.BackColor = SystemColors.Control;
        }

        private void toolStripButton13_Click(object sender, EventArgs e)
        {
            rtxGhi.SelectionAlignment = HorizontalAlignment.Right;
            tsblef.BackColor = SystemColors.Control;
            tsbCenter.BackColor = SystemColors.Control;
            tsbright.BackColor = SystemColors.GradientActiveCaption;
        }

        private void toolStripButton10_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                   rtxGhi.ForeColor = colorDialog.Color;
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                    rtxGhi.BackColor = colorDialog.Color;
            }
        }
        private bool checkbold = false;
        private bool checkItalic = false;
        private bool checkunderline = false;
        private bool checkStrikeout = false;

        private void btnFont_Click(object sender, EventArgs e)
        {
            try
            {
                ToolStripButton s = (ToolStripButton)sender;
                if (s.Name == "tsbBold")
                {
                    if (s.BackColor == SystemColors.Control)
                    {
                        rtxGhi.SelectionFont = new Font(fontChu,sizeChu, (FontStyle.Bold |
                            (checkItalic ? FontStyle.Italic : 0) |
                            (checkunderline ? FontStyle.Underline : 0) |
                            (checkStrikeout ? FontStyle.Strikeout : 0)));
                        s.BackColor = SystemColors.GradientActiveCaption;
                        checkbold = true;
                    }
                    else
                    {
                        rtxGhi.SelectionFont = new Font(fontChu, sizeChu, (FontStyle.Regular |
                            (checkItalic ? FontStyle.Italic : 0) |
                            (checkunderline ? FontStyle.Underline : 0) |
                            (checkStrikeout ? FontStyle.Strikeout : 0)));
                        s.BackColor = SystemColors.Control;
                        checkbold = false;
                    }
                }
                else if (s.Name == "tsbItalic")
                    {
                        if (s.BackColor == SystemColors.Control)
                        {
                            rtxGhi.SelectionFont = new Font(fontChu, sizeChu, (FontStyle.Italic |
                                (checkbold ? FontStyle.Bold : 0) |
                                (checkunderline ? FontStyle.Underline : 0) |
                                (checkStrikeout ? FontStyle.Strikeout : 0)));
                            s.BackColor = SystemColors.GradientActiveCaption;
                            checkItalic = true;
                        }
                        else
                        {
                            rtxGhi.SelectionFont = new Font(fontChu, sizeChu, (FontStyle.Regular |
                                (checkbold ? FontStyle.Bold : 0) |
                                (checkunderline ? FontStyle.Underline : 0) |
                                (checkStrikeout ? FontStyle.Strikeout : 0)));
                            s.BackColor = SystemColors.Control;
                            checkItalic = false;
                        }
                    }
                else if (s.Name == "tsbUnderline")
                {
                    if (s.BackColor == SystemColors.Control)
                    {
                        rtxGhi.SelectionFont = new Font(fontChu, sizeChu, (FontStyle.Underline |
                            (checkItalic ? FontStyle.Italic : 0) |
                            (checkbold ? FontStyle.Bold : 0) |
                            (checkStrikeout ? FontStyle.Strikeout : 0)));
                        s.BackColor = SystemColors.GradientActiveCaption;
                        checkunderline = true;
                    }
                    else
                    {
                        rtxGhi.SelectionFont = new Font(fontChu, sizeChu, (FontStyle.Regular |
                            (checkItalic ? FontStyle.Italic : 0) |
                            (checkbold ? FontStyle.Bold : 0) |
                            (checkStrikeout ? FontStyle.Strikeout : 0)));
                        s.BackColor = SystemColors.Control;
                        checkunderline = false;
                    }
                }
                else if (s.Name == "tsbStrikeout")
                {
                    if (s.BackColor == SystemColors.Control)
                    {
                        rtxGhi.SelectionFont = new Font(fontChu, sizeChu, (FontStyle.Strikeout |
                            (checkbold ? FontStyle.Bold : 0) |
                            (checkunderline ? FontStyle.Underline : 0) |
                            (checkItalic ? FontStyle.Italic : 0)));
                        s.BackColor = SystemColors.GradientActiveCaption;
                        checkStrikeout = true;
                    }
                    else
                    {
                        rtxGhi.SelectionFont = new Font(fontChu, sizeChu, (FontStyle.Regular |
                            (checkbold ? FontStyle.Bold : 0) |
                            (checkunderline ? FontStyle.Underline : 0) |
                            (checkItalic ? FontStyle.Italic : 0)));
                        s.BackColor = SystemColors.Control;
                        checkStrikeout = false;
                    }
                }
            }
            catch
            {

            }
        }

        private void toolStripButton14_Click(object sender, EventArgs e)
        {
            frmTK timkiem = new frmTK(rtxGhi);
            timkiem.Show(this);
        }

        private void toolStripButton15_Click(object sender, EventArgs e)
        {
            ThayThe thayThe = new ThayThe(rtxGhi);
            thayThe.Show(this);
        }

        private float sizeChu = 11;
        private string fontChu = "Times New Roman";

        // font chữ
        private void tscmbFont_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                fontChu = tscmbFont.SelectedItem.ToString();
                rtxGhi.SelectionFont = new Font(fontChu, sizeChu);
            }
            catch
            {
            }
        }

        private void tscFontsize_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                sizeChu = float.Parse(tscFontsize.SelectedItem.ToString());
                rtxGhi.SelectionFont = new Font(fontChu, sizeChu);
            }
            catch
            {

            }
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            
            if(tscFontsize.Items.Count > tscFontsize.SelectedIndex+1)
            {
                tscFontsize.SelectedIndex += 1;
            }    
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            if (  tscFontsize.SelectedIndex - 1 > 0)
            {
                tscFontsize.SelectedIndex -= 1;
            }
        }

        private void toolStripButton10_Click_1(object sender, EventArgs e)
        {
            ColorDialog color = new ColorDialog();
            if (color.ShowDialog() == DialogResult.OK)
            {
                rtxGhi.SelectionColor = color.Color;
                lbColor.BackColor = color.Color;
            }

        }

        private void toolStripButton1_Click_1(object sender, EventArgs e)
        {
            ColorDialog color = new ColorDialog();
            if (color.ShowDialog() == DialogResult.OK)
            {
                rtxGhi.SelectionBackColor = color.Color;
                lbBcolor.BackColor= color.Color;
            }

        }

        private void toolStripButton11_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Nút dự phòng k có sự kiện");
        }

        private void toolStripButton17_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog insertImg = new OpenFileDialog();
                insertImg.Filter = "Image|*.bmp;*.jpg;*.gif;*.png;*.tif";
                insertImg.ShowDialog();
                string path = insertImg.FileName;
                if (path != "")
                {
                    Clipboard.SetImage(Image.FromFile(path));
                    rtxGhi.Paste();
                }
            }
            catch
            {
            }
        }

        private void chenEmoji()
        {
            string duongDan = Environment.CurrentDirectory.ToString(); // lấy đường dẫn thư mục
            var url = Directory.GetParent(Directory.GetParent(duongDan).ToString()); // lấy thư mục cha
            string path = url + @"\emoji"; // lấy đường dẫn

            string[] files = Directory.GetFiles(path); // lấy tên file là ảnh

            foreach (String f in files)
            {
                Image img = Image.FromFile(f);  // từ cái file đó chuyển qua định dạng ảnh
                imageList1.Images.Add(img); // bỏ vào img list
            }
            this.listView1.View = View.LargeIcon; // thuộc tính ảnh to hay nhỏ
            this.imageList1.ImageSize = new Size(32, 32); //size

            this.listView1.LargeImageList = this.imageList1; // ép thuộc tính vào listvieww
            // ép ảnh vào list vieww
            for (int i = 0; i < this.imageList1.Images.Count; i++)
            {
                this.listView1.Items.Add(" ", i);
            }

        }
        private bool check = true;
        private void toolStripButton16_Click(object sender, EventArgs e)
        {
            if(check)
            {
                listView1.Visible = check;
                check = false;
            }
            else
            {
                listView1.Visible = check;
                check = true;
            }
            
        }

        private int id = 0;
        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (listView1.SelectedIndices.Count <= 0) return;
                if (listView1.FocusedItem == null) return;
                id = listView1.SelectedIndices[0];
                if (id < 0) return; // nếu mà id = 0 tức là ảnh lỗi hoặc chưa có ảnh nên cho qua
                Clipboard.SetImage(imageList1.Images[id]); // dán icon vào 
                rtxGhi.Paste(); // dán vào n
            }
            catch
            {

            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (checksave == true)
            {
                DialogResult savePrompt = MessageBox.Show("Do you want to save changes to Document?", "SAVE", MessageBoxButtons.YesNoCancel);
                switch (savePrompt)
                {
                    case DialogResult.Cancel:
                        break;
                    case DialogResult.No:
                        Close();
                        break;
                    case DialogResult.Yes:
                        saveAsToolStripMenuItem.PerformClick();
                        this.Close();
                        break;
                }
            }
            else
            {
                this.Close();
            }
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (checksave == true)
            {
                DialogResult savePrompt = MessageBox.Show("Do you want to save changes to Document?", "SAVE", MessageBoxButtons.YesNoCancel);
                switch (savePrompt)
                {
                    case DialogResult.Cancel:
                        break;
                    case DialogResult.No:
                        panel7.Visible = true;
                        rtxGhi.Visible = true;
                        rtxGhi.Text = "";
                        tscmbFont.SelectedItem = "Times New Roman";
                        tscFontsize.SelectedItem = "11";
                        lbtenfile.Text = "New File";
                        LayFileHienTai = "";
                        FileHienTai = false;
                        break;
                    case DialogResult.Yes:
                        saveAsToolStripMenuItem.PerformClick();
                        panel7.Visible = true;
                        rtxGhi.Visible = true;
                        rtxGhi.Text = "";
                        tscmbFont.SelectedItem = "Times New Roman";
                        tscFontsize.SelectedItem = "11";
                        lbtenfile.Text = "New File";
                        LayFileHienTai = "";
                        FileHienTai = false;
                        break;
                }
            }
            else
            {
                panel7.Visible = true;
                rtxGhi.Visible = true;
                rtxGhi.Text = "";
                tscmbFont.SelectedItem = "Times New Roman";
                tscFontsize.SelectedItem = "11";
                lbtenfile.Text = "New File";
                LayFileHienTai = "";
                FileHienTai = false;
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(FileHienTai)
            {
               if(LayFileHienTai != "") rtxGhi.SaveFile(LayFileHienTai);
            }
            else
            {
                saveAsToolStripMenuItem.PerformClick();
            }
            checksave = true;
        }
        private string LayFileHienTai = "";
        private bool FileHienTai = false;
        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog s = new SaveFileDialog();
                s.Filter = "Save File (*.rtf)|*.rtf| (*.txt)|*.txt";
                if (s.ShowDialog() == DialogResult.OK)
                {
                    string path = s.FileName;
                    if (path != "")
                    {
                        rtxGhi.SaveFile(path);
                        LayFileHienTai = path;
                        FileHienTai = true;
                        tenfile(path);
                    }
                }
                checksave = true;
            }
            catch
            {
            }
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (checksave == true)
                {
                    DialogResult savePrompt = MessageBox.Show("Do you want to save changes to Document?", "SAVE", MessageBoxButtons.YesNoCancel);
                    switch (savePrompt)
                    {
                        case DialogResult.Cancel:
                            break;
                        case DialogResult.No:
                            rtxGhi.Text = "";
                            lbtenfile.Text = "";
                            LayFileHienTai = "";
                            FileHienTai = false;
                            rtxGhi.Visible = false;
                            panel7.Visible = false;
                            break;
                        case DialogResult.Yes:
                            saveAsToolStripMenuItem.PerformClick();
                            rtxGhi.Text = "";
                            lbtenfile.Text = "";
                            LayFileHienTai = "";
                            FileHienTai = false;
                            rtxGhi.Visible = false;
                            panel7.Visible = false;
                            break;
                    }
                }
                else
                {
                    rtxGhi.Text = "";
                    lbtenfile.Text = "";
                    LayFileHienTai = "";
                    FileHienTai = false;
                    rtxGhi.Visible = false;
                    panel7.Visible = false;
                }
                
            }
            catch
            {
            }
        }

        private void rtxGhi_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (rtxGhi.SelectionFont != null)
                {
                    lbBcolor.BackColor = rtxGhi.SelectionBackColor;
                    lbColor.BackColor = rtxGhi.SelectionColor;

                    tsbBold.BackColor = SystemColors.Control;
                    tsbItalic.BackColor = SystemColors.Control;
                    tsbUnderline.BackColor = SystemColors.Control;
                    tsbStrikeout.BackColor = SystemColors.Control;
                    if (rtxGhi.SelectionFont.Bold)
                    {
                        tsbBold.BackColor = SystemColors.GradientActiveCaption;
                    }
                    if (rtxGhi.SelectionFont.Italic)
                    {
                        tsbItalic.BackColor = SystemColors.GradientActiveCaption;
                    }
                    if (rtxGhi.SelectionFont.Underline)
                    {
                        tsbUnderline.BackColor = SystemColors.GradientActiveCaption;
                    }
                    if (rtxGhi.SelectionFont.Strikeout)
                    {
                        tsbStrikeout.BackColor = SystemColors.GradientActiveCaption;
                    }
                    tsblef.BackColor = SystemColors.Control;
                    tsbCenter.BackColor = SystemColors.Control;
                    tsbright.BackColor = SystemColors.Control;
                    switch (rtxGhi.SelectionAlignment)
                    {
                        case HorizontalAlignment.Left:
                            tsblef.BackColor = SystemColors.GradientActiveCaption;
                            break;
                        case HorizontalAlignment.Center:
                            tsbCenter.BackColor = SystemColors.GradientActiveCaption;
                            break;
                        case HorizontalAlignment.Right:
                            tsbright.BackColor = SystemColors.GradientActiveCaption;
                            break;
                    }
                }
            }
            catch
            {
            }
        }

        private void toolStripButton18_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Nút dự phòng k có sự kiện");
        }

        private void toolStripButton19_Click(object sender, EventArgs e)
        {
           this.rtxGhi.Undo();
        }

        private void toolStrip4_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            this.rtxGhi.Redo();
        }

        private void rtxGhi_TextChanged(object sender, EventArgs e)
        {
            checksave = true;
        }

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (printDialog1.ShowDialog() == DialogResult.OK)
                {
                    printDocument1.Print();
                }
            }
            catch
            {
            }
      
        }
       
        private int linesPrinted;
        private string[] lines;
        private void printDocument1_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            try
            {
                char[] param = { '\n' };
                if (printDocument1.PrinterSettings.PrintRange == PrintRange.Selection)
                {
                    lines = rtxGhi.SelectedText.Split(param);
                }
                else
                {
                    lines = rtxGhi.Text.Split(param);
                }

                int i = 0;
                char[] trimParam = { '\r' };
                foreach (string s in lines)
                {
                    lines[i++] = s.TrimEnd(trimParam);
                }
            }
            catch
            {
            }
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
                try
                {
                    if (lines.Length <= 0) return;
                    int x = e.MarginBounds.Left;
                    int y = e.MarginBounds.Top;
                    Brush brush = new SolidBrush(rtxGhi.ForeColor);

                    while (linesPrinted < lines.Length)
                    {
                        e.Graphics.DrawString(lines[linesPrinted++],
                            rtxGhi.Font, brush, x, y);
                        y += 15;
                        if (y >= e.MarginBounds.Bottom)
                        {
                            e.HasMorePages = true;
                            return;
                        }
                    }
                    linesPrinted = 0;
                    e.HasMorePages = false;
                }
                catch
                {
                }
            }
    }
}
