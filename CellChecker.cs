using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using System.IO.Compression;

namespace DataAnalysisTool
{
    public partial class DataAnalysisTool
    {
        //------------------CELL LENGTH CHECKER START------------------------------------------------------

        private void cellLength_Click(object sender, EventArgs e)
        {
            int a = 0;
            String reqItem;
            String specialChar = textBox1.Text;
            if (textBox1.Text.Length == 0)
            {
                MessageBox.Show("You did not select a special character!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            foreach (Object selecteditem in specialCharacterCheckerListBox.SelectedItems)
            {
                a++;
                reqItem = selecteditem as String;
                int specialCharacterCurIndex = specialCharacterCheckerListBox.Items.IndexOf(reqItem);
                if (specialCharacterCurIndex >= 0)
                {

                    for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                    {

                        var value = importedfileDataGridView.Rows[i].Cells[specialCharacterCurIndex].Value.ToString();
                        //MessageBox.Show("value "+value+"reqitem "+reqItem);
                        if (value.Contains(specialChar) == true)
                        {
                            MessageBox.Show("'" + specialChar + "'" + " WAS found in the column " + "'" + selecteditem + "'", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            return;
                        }
                    }
                }
            }
            if (a == 0)
            {
                MessageBox.Show("You did not select a column!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            MessageBox.Show("'" + specialChar + "'" + " WAS NOT FOUND!", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);










            //{
            //    try
            //    {
            //        DataGridViewColumn column = importedfileDataGridView.Columns[textBox3.Text];
            //        MessageBox.Show(column.Name + " must be " + textBox4.Text + " Digit(s) Long!");
            //    }
            //    catch (Exception ex)
            //    {
            //        if (textBox3.Text.Length == 0)
            //        {
            //            MessageBox.Show("You did not enter a column name!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            //            return;
            //        }
            //        if (textBox4.Text.Length == 0)
            //        {
            //            MessageBox.Show("You did not enter a length!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
            //            return;
            //        }
            //        MessageBox.Show(ex.Message, "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    }
            //}
        }

        //------------------CELL LENGTH CHECKER END------------------------------------------------------

        //------------------SPECIAL CHARACTER CHECKER START------------------------------------------------------

        private void specialCharacter_Click(object sender, EventArgs e)
        {
            int a = 0;
            String reqItem;
            String specialChar=textBox1.Text;
            if (textBox1.Text.Length == 0)
            {
                MessageBox.Show("You did not select a special character!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            foreach (Object selecteditem in specialCharacterCheckerListBox.SelectedItems)
            {
                a++;
                reqItem = selecteditem as String;
                int specialCharacterCurIndex = specialCharacterCheckerListBox.Items.IndexOf(reqItem);
                if (specialCharacterCurIndex >= 0)
                {

                    for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                    {
                        
                        var value = importedfileDataGridView.Rows[i].Cells[specialCharacterCurIndex].Value.ToString();
                        //MessageBox.Show("value "+value+"reqitem "+reqItem);
                        if (value.Contains(specialChar) == true)
                        {
                            MessageBox.Show("'" + specialChar + "'" + " WAS found in the column " + "'" + selecteditem + "'", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            return;
                        }
                    }
                }
            }
            if (a == 0)
            {
                MessageBox.Show("You did not select a column!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            MessageBox.Show("'" + specialChar + "'" + " WAS NOT FOUND!", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
        }

        //------------------SPECIAL CHARACTER CHECKER END------------------------------------------------------

        //------------------NULL CHECKER START------------------------------------------------------

        private void nullChecker_Click(object sender, EventArgs e)
        {
            int a = 0;
            String reqItem;
            foreach (Object selecteditem in nullCheckerListBox.SelectedItems)
            {
                a++;
                reqItem = selecteditem as String;
                int nullCheckCurIndex = nullCheckerListBox.Items.IndexOf(reqItem);
                if (nullCheckCurIndex >= 0)
                {

                    for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                    {

                        var value = importedfileDataGridView.Rows[i].Cells[nullCheckCurIndex].Value.ToString();
                        if (string.IsNullOrWhiteSpace(value))
                        {
                            MessageBox.Show("NULL value found in column " + "'" + reqItem + "'" + " at line " + (i + 1), "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            
                            return;
                        }
                    }
                }
            }
            if (a == 0)
            {
                MessageBox.Show("You did not select a column!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            MessageBox.Show("no NULL value!", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
        }
        //------------------NULL CHECKER END------------------------------------------------------

    }
}
