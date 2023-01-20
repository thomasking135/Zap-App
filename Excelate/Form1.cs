using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excelate
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            MeExit();
        }
        private void MeExit()
        {
            DialogResult iExit;
            iExit = MessageBox.Show("Confirm if you want to exit", "Save Zap file", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (iExit == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        private void exitToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            MeExit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add(txtStudent_ID.Text, txtFirstname.Text, txtLastname.Text, txtAddress.Text, txtDOB.Text, txtMobile.Text);
            clear();
        }

        private void addNewToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add(txtStudent_ID.Text, txtFirstname.Text, txtLastname.Text, txtAddress.Text, txtDOB.Text, txtMobile.Text);
            clear();
        }

        private void iDelete()
        {
            foreach (DataGridViewRow item in this.dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.RemoveAt(item.Index);
                
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            iDelete();
        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            iDelete();
        }

        private void iReset()
        {
        //clears text from the form
            foreach (var c in this.Controls)
            {
                if (c is TextBox)
                {
                    ((TextBox)c).Text = String.Empty;
                }
            }

        //clears data from the data grid
            int numRows = dataGridView1.Rows.Count;
            for (int i = 0; i < numRows; i++)
            {
                try
                {
                    int max = dataGridView1.Rows.Count - 1;
                    dataGridView1.Rows.Remove(dataGridView1.Rows[max]);
                }
                catch (Exception exe)
                {
                    MessageBox.Show("All rows are to be deleted " + exe, "Zap Delete",
                        MessageBoxButtons.OK, MessageBoxIcon.Information); 
                }
            }

            
        }
        private void button2_Click(object sender, EventArgs e)
        {
            iReset();
        }

        private void resetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            iReset();
        }

        Bitmap bitmap;
        private void button3_Click(object sender, EventArgs e)
        {
            int height = dataGridView1.Height;
            dataGridView1.Height = dataGridView1.RowCount * dataGridView1.RowTemplate.Height * 2;
            bitmap = new Bitmap(dataGridView1.Width, dataGridView1.Height);
            dataGridView1.DrawToBitmap(bitmap, new Rectangle(0, 0, dataGridView1.Width, dataGridView1.Height));
            printPreviewDialog1.PrintPreviewControl.Zoom = 1;
            printPreviewDialog1.ShowDialog();
            dataGridView1.Height = height;
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawImage(bitmap, 0, 0);
        }

        private void printToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            int height = dataGridView1.Height;
            dataGridView1.Height = dataGridView1.RowCount * dataGridView1.RowTemplate.Height * 2;
            bitmap = new Bitmap(dataGridView1.Width, dataGridView1.Height);
            dataGridView1.DrawToBitmap(bitmap, new Rectangle(0, 0, dataGridView1.Width, dataGridView1.Height));
            printPreviewDialog1.PrintPreviewControl.Zoom = 1;
            printPreviewDialog1.ShowDialog();
            dataGridView1.Height = height;
        }


        private void iSave()
        {

            //clears text from the form
            


            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = null;

            app.Visible = true;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "Exported from Zap";

            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                
            }
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    
                }
            }
            foreach (var c in this.Controls)
            {
                if (c is TextBox)
                {
                    ((TextBox)c).Text = String.Empty;
                    
                }
            }

            
        }

        private void button5_Click(object sender, EventArgs e)
        {
            iSave();
        }

        private void saveToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            iSave();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //Form
        }

        private void clear()
        {
            txtStudent_ID.Clear();
            txtFirstname.Clear();
            txtLastname.Clear();
            txtAddress.Clear();
            txtDOB.Clear();
            txtMobile.Clear();
        }
        private void button7_Click(object sender, EventArgs e)
        {
            txtStudent_ID.Clear();
            txtFirstname.Clear();
            txtLastname.Clear();
            txtAddress.Clear();
            txtDOB.Clear();
            txtMobile.Clear();
        }
    }
}


