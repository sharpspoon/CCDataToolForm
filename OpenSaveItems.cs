using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using System.IO.Compression;

namespace CCDataImportTool
{
    public partial class CCDataTool
    {

        //------------------OPEN/SAVE CSV START------------------------------------------------------
        private void menu_Open_Csv_Click(object sender, EventArgs e)
        {
            try
            {
                    
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "CSV | *.csv", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                        dataGridView1.DataSource = ReadCsv(ofd.FileName);
                    toolStripStatusLabel13.Text = ofd.FileName;
                    toolStripStatusLabel4.Text = dataGridView1.Rows.Count.ToString();
                    richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading CSV: " + ofd.FileName + "...Done.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        public DataTable ReadCsv(string fileName)
        {
            DataTable dt = new DataTable("Data");
            using (OleDbConnection cn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\"" +
                Path.GetDirectoryName(fileName) + "\";Extended Properties='text;HDR=yes;FMT=Delimited(,)';"))
            {
                using (OleDbCommand cmd = new OleDbCommand(string.Format("select * from [{0}]", new FileInfo(fileName).Name), cn))
                {
                    cn.Open();
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
            }
            
            return dt;

        }
        protected void menu_Save_Csv_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "CSV|*.csv";
            if (this.saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // create one file gridview.csv in writing mode using streamwriter
                StreamWriter sw = new StreamWriter("c:\\gridview.csv");
                // now add the gridview header in csv file suffix with "," delimeter except last one
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    sw.Write(dataGridView1.Columns[i].HeaderText);
                    if (i != dataGridView1.Columns.Count)
                    {
                        sw.Write(",");
                    }
                }
                // add new line
                sw.Write(sw.NewLine);
                // iterate through all the rows within the gridview
                foreach (DataGridViewRow dr in dataGridView1.Rows)
                {
                    // iterate through all colums of specific row
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        // write particular cell to csv file
                        sw.Write(dr.Cells[i]);
                        if (i != dataGridView1.Columns.Count)
                        {
                            sw.Write(",");
                        }
                    }
                    // write new line
                    sw.Write(sw.NewLine);
                }
                // flush from the buffers.
                sw.Flush();
                // closes the file
                sw.Close();
            }
        }

        //------------------OPEN/SAVE CSV END------------------------------------------------------

        //------------------OPEN/SAVE XML START------------------------------------------------------

        private void menu_Open_Xml_Click(object sender, EventArgs e)
        {
            try
            {
                DataSet dataSet = new DataSet();
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "XML | *.xml", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                        dataSet.ReadXml(ofd.FileName);
                    dataGridView1.DataSource = dataSet.Tables[0];

                    toolStripStatusLabel13.Text = ofd.FileName;
                    toolStripStatusLabel4.Text = dataGridView1[0, dataGridView1.Rows.Count - 1].Value.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void menu_Save_Xml_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "XML|*.xml";
            if (this.saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                DataTable dt = (DataTable)this.dataGridView1.DataSource;
                dt.WriteXml(this.saveFileDialog1.FileName, XmlWriteMode.WriteSchema);
            }
        }

        //------------------OPEN/SAVE XML END------------------------------------------------------

        //------------------OPEN/SAVE XLS START------------------------------------------------------

        private void menu_Open_Xls_Click(object sender, EventArgs e)
        {
            try
            {

                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "XLS | *.xls", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                        dataGridView1.DataSource = ReadXls(ofd.FileName);
                    toolStripStatusLabel13.Text = ofd.FileName;
                    toolStripStatusLabel4.Text = dataGridView1.Rows.Count.ToString();
                    richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading XLS: " + ofd.FileName + "...Done.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public DataTable ReadXls(string fileName)
        {
            DataTable dt = new DataTable();
            using (OleDbConnection cn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\"" +
                Path.GetDirectoryName(fileName) + "\";Extended Properties='Excel 8.0;HDR=YES;';"))
            {
                using (OleDbCommand cmd = new OleDbCommand("select * from [" + fileName + "$]", cn))
                {
                    cn.Open();
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
            }

            return dt;

        }
        //------------------OPEN/SAVE XLS END------------------------------------------------------

        //------------------PRINT DOCUMENT START------------------------------------------------------

        Bitmap bitmap;
        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0 || dataGridView1.Rows == null)
            {
                MessageBox.Show("No data to print", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                //Resize DataGridView to full height.
                int height = dataGridView1.Height;
                dataGridView1.Height = dataGridView1.RowCount * dataGridView1.RowTemplate.Height;

                //Create a Bitmap and draw the DataGridView on it.
                bitmap = new Bitmap(this.dataGridView1.Width, this.dataGridView1.Height);
                dataGridView1.DrawToBitmap(bitmap, new Rectangle(0, 0, this.dataGridView1.Width, this.dataGridView1.Height));

                //Resize DataGridView back to original height.
                dataGridView1.Height = height;

                //Show the Print Preview Dialog.
                printPreviewDialog1.Document = printDocument1;
                printPreviewDialog1.PrintPreviewControl.Zoom = 1;
                printPreviewDialog1.ShowDialog();
            }
        }
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            //Print the contents.
            e.Graphics.DrawImage(bitmap, 0, 0);
        }

        //------------------PRINT DOCUMENT END------------------------------------------------------

    }
}
