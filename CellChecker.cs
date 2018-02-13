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
        //------------------CELL LENGTH CHECKER START------------------------------------------------------

        private void cellLength_Click(object sender, EventArgs e)
        {
            {
                try
                {
                    DataGridViewColumn column = dataGridView1.Columns[textBox3.Text];
                    MessageBox.Show(column.Name + " must be " + textBox4.Text + " Digit(s) Long!");
                }
                catch (Exception ex)
                {
                    if (textBox3.Text.Length == 0)
                    {
                        MessageBox.Show("You did not enter a column name!\r\nThe operation will now cancel.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }
                    if (textBox4.Text.Length == 0)
                    {
                        MessageBox.Show("You did not enter a length!\r\nThe operation will now cancel.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }
                    MessageBox.Show(ex.Message, "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        //------------------CELL LENGTH CHECKER END------------------------------------------------------

        //------------------SPECIAL CHARACTER CHECKER START------------------------------------------------------

        private void specialCharacter_Click(object sender, EventArgs e)
        {
            String searchValue = comboBox1.Text;
            string specialBoxFill = textBox5.Text;
            if (textBox5.Text.Length == 0)
            {
                MessageBox.Show("You did not enter a column name!\r\nThe operation will now cancel.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            if (comboBox1.Text.Length == 0)
            {
                MessageBox.Show("You did not select a special character!\r\nThe operation will now cancel.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    if (row.Cells[textBox5.Text].Value.ToString().Contains(comboBox1.Text))
                    {
                        MessageBox.Show("'" + comboBox1.Text + "'" + " WAS found in the column " + "'" + textBox5.Text + "'", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }
                }
                catch
                {
                    MessageBox.Show("'" + comboBox1.Text + "'" + " WAS NOT  found in column " + "'" + textBox5.Text + "'", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    return;
                }
            }
        }

        //------------------SPECIAL CHARACTER CHECKER END------------------------------------------------------

        //------------------NULL CHECKER START------------------------------------------------------

        private void nullChecker_Click(object sender, EventArgs e)
        {
            if (textBox6.Text.Length == 0)
            {
                MessageBox.Show("You did not enter a column name!\r\nThe operation will now cancel.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                try
                {
                    var value = dataGridView1.Rows[i].Cells[textBox6.Text].Value.ToString();
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        MessageBox.Show("NULL value found in column " + "'" + textBox6.Text + "'" + " at line " + dataGridView1.Rows[i + 1]);
                        return;
                    }
                }
                catch (Exception)
                {
                    // If we have reached this far, then none of the cells were empty.
                    MessageBox.Show("No NULL values found in column " + "'" + textBox6.Text + "'");
                    return;
                }
            }
        }
        private void textBox6_TextChanged(object sender, EventArgs e)
        {
        }

        //------------------NULL CHECKER END------------------------------------------------------

    }
}
