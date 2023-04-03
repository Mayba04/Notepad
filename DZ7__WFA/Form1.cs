using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace DZ7__WFA
{
    

    public partial class Form1 : Form
    {
        //public string File_Name { get; set; }
        private string A_Z;
        private string a_z;
        private string numbers;
        F_N File_Name;
        public Form1()
        {
            A_Z = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            a_z = "abcdefghijklmnopqrstuvwxyz";
            numbers = "123456789";
            File_Name.Index = -2;
            InitializeComponent();
            tabControl1.TabPages.Remove(tabPage2);
            //tabControl1.TabPages.Remove(tabPage1);
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "DOC File (.DOC)|*.doc |RTF File (.RTF)|*.rtf";
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //openFileDialog1.FilterIndex = 2;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var res = MessageBox.Show("Rewrite the page or create a new page?", "Overwriting the page", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (res == DialogResult.Yes)
                {
                    GetRichTextBox().LoadFile(openFileDialog1.FileName, RichTextBoxStreamType.RichText);
                    File_Name.Name = openFileDialog1.FileName;
                    File_Name.Index = tabControl1.SelectedIndex;
                    int a = tabControl1.SelectedIndex;
                    MessageBox.Show("File open!", "File", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    TabPage page = new TabPage($"{openFileDialog1.SafeFileName}");

                    RichTextBox richTextBox2 = new RichTextBox();
                    richTextBox2.ContextMenuStrip = this.contextMenuStrip1;
                    richTextBox2.Dock = System.Windows.Forms.DockStyle.Fill;
                    richTextBox2.Location = new System.Drawing.Point(3, 3);
                    richTextBox2.Name = "richTextBox2";
                    richTextBox2.Size = new System.Drawing.Size(786, 356);
                    richTextBox2.TabIndex = 3;
                    richTextBox2.Text = "";
                    this.toolTip1.SetToolTip(richTextBox2, "Write the txet");
                    richTextBox2.AllowDrop = true;
                    richTextBox2.DragDrop += RichTextBox1_DragDrop;
                    richTextBox2.TextChanged += new System.EventHandler(this.richTextBox1_TextChanged);
                    // 
                    richTextBox2.LoadFile(openFileDialog1.FileName, RichTextBoxStreamType.RichText);
                    page.Controls.Add(richTextBox2);

                    tabControl1.TabPages.Add(page);
                
                    File_Name.Name = openFileDialog1.FileName;
                    File_Name.Index = tabControl1.TabPages.Count - 1;
                    MessageBox.Show("File open!", "File", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            

            if (File_Name.Index == tabControl1.SelectedIndex)
            {
                GetRichTextBox().SaveFile(File_Name.Name, RichTextBoxStreamType.RichText);
                MessageBox.Show("File save!");
            }
            else
            {
                saveAsToolStripMenuItem_Click(sender, e);
            }
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.DefaultExt = ".doc";
            saveFileDialog1.Filter = "DOC File (.DOC)|*.doc |RTF File (.RTF)|*.rtf";
            saveFileDialog1.OverwritePrompt = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                GetRichTextBox().SaveFile(saveFileDialog1.FileName, RichTextBoxStreamType.RichText);
                MessageBox.Show("File save!");
            }
        }

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (printDialog1.ShowDialog() == DialogResult.OK)
                printDocument1.Print();
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //RichTextBox richTextBox = GetRichTextBox();
            //if (richTextBox.TextLength > 0)
            //{
            //    richTextBox.Copy();
            //}
            GetRichTextBox().Copy();
        }

        private RichTextBox GetRichTextBox()
        {
            RichTextBox richTextBox = new RichTextBox();
            TabPage page = tabControl1.SelectedTab;
            if (page != null)
            {
                richTextBox = page.Controls[0] as RichTextBox;
            }
            return richTextBox;
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GetRichTextBox().Paste();
        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GetRichTextBox().Cut();
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GetRichTextBox().SelectAll();
        }

        private void fontColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            colorDialog1.ShowDialog();
            GetRichTextBox().ForeColor = colorDialog1.Color;
            toolStripButton9.BackColor = colorDialog1.Color;
            //richTextBox1.ForeColor = colorDialog1.Color;
        }

        private void typeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            fontDialog1.ShowDialog();
            GetRichTextBox().Font = fontDialog1.Font;
        }

        private void backcolorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            colorDialog1.ShowDialog();
            GetRichTextBox().BackColor = colorDialog1.Color;
        }

        private void printDocument1_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            char[] param = { '\n' };

            if (printDialog1.PrinterSettings.PrintRange == PrintRange.Selection)
            {
                lines = GetRichTextBox().SelectedText.Split(param);
            }
            else
            {
                lines = GetRichTextBox().Text.Split(param);
            }

            int i = 0;
            char[] trimParam = { '\r' };
            foreach (string s in lines)
            {
                lines[i++] = s.TrimEnd(trimParam);
            }
        }

        private int linesPrinted;
        private string[] lines;

        private void OnPrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            int x = e.MarginBounds.Left;
            int y = e.MarginBounds.Top;
            Brush brush = new SolidBrush(GetRichTextBox().ForeColor);

            while (linesPrinted < lines.Length)
            {
                e.Graphics.DrawString(lines[linesPrinted++],
                     GetRichTextBox().Font, brush, x, y);
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

        private void appealToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            TabPage page = new TabPage($"New tab {tabControl1.TabPages.Count + 1}");
            // richTextBox1
            // 
            RichTextBox richTextBox2 = new RichTextBox();
            richTextBox2.ContextMenuStrip = this.contextMenuStrip1;
            richTextBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            richTextBox2.Location = new System.Drawing.Point(3, 3);
            richTextBox2.Name = "richTextBox2";
            richTextBox2.Size = new System.Drawing.Size(786, 356);
            richTextBox2.TabIndex = 3;
            richTextBox2.Text = "";
            this.toolTip1.SetToolTip(richTextBox2, "Write the txet");
            richTextBox2.AllowDrop = true;
            richTextBox2.DragDrop += RichTextBox1_DragDrop;
            richTextBox2.TextChanged += new System.EventHandler(this.richTextBox1_TextChanged);
            // 
            page.Controls.Add(richTextBox2);

            tabControl1.TabPages.Add(page);
        }

        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex != -1)
            {
                tabControl1.TabPages.RemoveAt(tabControl1.SelectedIndex);
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            int tmp_letters = 0;
            int tmp_numbers = 0;
            string text = GetRichTextBox().Text;
            string[] words = text.Split(new char[] { ' ', '\n' });
            toolStripStatusLabel2.Text = words.Length.ToString();

            for (int i = 0; i < text.Length; i++)
            {
                for (int j = 0; j < A_Z.Length; j++)
                {
                    if (text[i] == A_Z[j] || text[i] == a_z[j])
                    {
                        tmp_letters++;
                    }
                }

                for (int j = 0; j < numbers.Length; j++)
                {
                    if (text[i] == numbers[j])
                    {
                        tmp_numbers++;
                    }
                }
            }
            toolStripStatusLabel4.Text = tmp_letters.ToString();
            toolStripStatusLabel6.Text = tmp_numbers.ToString();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            richTextBox1.AllowDrop = true;
            richTextBox1.DragDrop += RichTextBox1_DragDrop;
        }

        private void RichTextBox1_DragDrop(object sender, DragEventArgs e)
        {
           
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var fileNames = e.Data.GetData(DataFormats.FileDrop) as string[];
                foreach (var item in fileNames)
                {
                    RichTextBox richTextBox = new RichTextBox();
                    richTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
                    richTextBox.LoadFile(item);
                    GetRichTextBox().AppendText(richTextBox.Text + "\n");
                }
            }
            else if (e.Data.GetDataPresent(DataFormats.Text))
            {
                string item = e.Data.GetData(DataFormats.Text).ToString();
                GetRichTextBox().AppendText(item + "\n");
            }
        }

        private void clearAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GetRichTextBox().Clear();
        }

        private void toolStripButton9_Click(object sender, EventArgs e)
        {

        }
    }

    struct F_N
    {
        public string Name { get; set; }
        public int Index { get; set; }
    }
}
