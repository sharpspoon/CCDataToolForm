﻿using System;
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

        //------------------MEDICARE CHECKER START------------------------------------------------------

        private void medicareButton_Click(object sender, EventArgs e)
        {

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                try
                {
                    if (dataGridView1.ColumnCount != 37)
                    {
                        MessageBox.Show("Medicare files need 37 columns. You have " + dataGridView1.ColumnCount + ".", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }
                    var value0 = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    var value1 = dataGridView1.Rows[i].Cells[1].Value.ToString();
                    var value2 = dataGridView1.Rows[i].Cells[2].Value.ToString();
                    var value3 = dataGridView1.Rows[i].Cells[3].Value.ToString();
                    var value4 = dataGridView1.Rows[i].Cells[4].Value.ToString();
                    var value5 = dataGridView1.Rows[i].Cells[5].Value.ToString();
                    var value6 = dataGridView1.Rows[i].Cells[6].Value.ToString();
                    var value7 = dataGridView1.Rows[i].Cells[7].Value.ToString();
                    var value8 = dataGridView1.Rows[i].Cells[8].Value.ToString();
                    var value9 = dataGridView1.Rows[i].Cells[9].Value.ToString();
                    var value10 = dataGridView1.Rows[i].Cells[10].Value.ToString();
                    var value11 = dataGridView1.Rows[i].Cells[11].Value.ToString();
                    var value12 = dataGridView1.Rows[i].Cells[12].Value.ToString();
                    var value13 = dataGridView1.Rows[i].Cells[13].Value.ToString();
                    var value14 = dataGridView1.Rows[i].Cells[14].Value.ToString();
                    var value15 = dataGridView1.Rows[i].Cells[15].Value.ToString();
                    var value16 = dataGridView1.Rows[i].Cells[16].Value.ToString();
                    var value17 = dataGridView1.Rows[i].Cells[17].Value.ToString();
                    var value18 = dataGridView1.Rows[i].Cells[18].Value.ToString();
                    var value19 = dataGridView1.Rows[i].Cells[19].Value.ToString();
                    var value20 = dataGridView1.Rows[i].Cells[20].Value.ToString();
                    var value21 = dataGridView1.Rows[i].Cells[21].Value.ToString();
                    var value22 = dataGridView1.Rows[i].Cells[22].Value.ToString();
                    var value23 = dataGridView1.Rows[i].Cells[23].Value.ToString();
                    var value24 = dataGridView1.Rows[i].Cells[24].Value.ToString();
                    var value25 = dataGridView1.Rows[i].Cells[25].Value.ToString();
                    var value26 = dataGridView1.Rows[i].Cells[26].Value.ToString();
                    var value27 = dataGridView1.Rows[i].Cells[27].Value.ToString();
                    var value28 = dataGridView1.Rows[i].Cells[28].Value.ToString();
                    var value29 = dataGridView1.Rows[i].Cells[29].Value.ToString();
                    var value30 = dataGridView1.Rows[i].Cells[30].Value.ToString();
                    var value31 = dataGridView1.Rows[i].Cells[31].Value.ToString();
                    var value32 = dataGridView1.Rows[i].Cells[32].Value.ToString();
                    var value33 = dataGridView1.Rows[i].Cells[33].Value.ToString();
                    var value34 = dataGridView1.Rows[i].Cells[34].Value.ToString();
                    var value35 = dataGridView1.Rows[i].Cells[35].Value.ToString();
                    var value36 = dataGridView1.Rows[i].Cells[36].Value.ToString();

                    //Required/Optional Check
                    if (string.IsNullOrWhiteSpace(value0))
                    {
                        MessageBox.Show("NULL value found in column #1 (CustomerId)  at line " + (i + 1) + " This is a required field.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(value1))
                    {
                        MessageBox.Show("NULL value found in column #2 (ContractNbr)  at line " + (i + 1) + " This is a required field.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(value2))
                    {
                        MessageBox.Show("NULL value found in column #3 (PBP)  at line " + (i + 1) + " This is a required field.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(value3))
                    {
                        MessageBox.Show("NULL value found in column #4 (HICN)  at line " + (i + 1) + " This is a required field.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(value6))
                    {
                        MessageBox.Show("NULL value found in column #7 (DatEff)  at line " + (i + 1) + " This is a required field.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(value8))
                    {
                        MessageBox.Show("NULL value found in column #9 (AppSignedDate)  at line " + (i + 1) + " This is a required field.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(value10))
                    {
                        MessageBox.Show("NULL value found in column #11 (Holder)  at line " + (i + 1) + " This is a required field.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(value23))
                    {
                        MessageBox.Show("NULL value found in column #24 (PolState)  at line " + (i + 1) + " This is a required field.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }


                    //Field Length Check
                    if (value0.Length > 30)
                    {
                        MessageBox.Show("column #1 (CustomerId)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value1.Length > 10)
                    {
                        MessageBox.Show("column #2 (ContractNbr)  needs to be 10 or less characters.  At line " + (i + 1) + " you have a value that is " + value1.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value2.Length > 10)
                    {
                        MessageBox.Show("column #3 (PBP)  needs to be 10 or less characters.  At line " + (i + 1) + " you have a value that is " + value2.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value3.Length > 20)
                    {
                        MessageBox.Show("column #4 (HICN)  needs to be 20 or less characters.  At line " + (i + 1) + " you have a value that is " + value3.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value4.Length > 30)
                    {
                        MessageBox.Show("column #5 (OED)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value4.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value5.Length > 30)
                    {
                        MessageBox.Show("column #6 (CMSOED)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value5.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value6.Length > 30)
                    {
                        MessageBox.Show("column #7 (DatEff)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value6.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value7.Length > 30)
                    {
                        MessageBox.Show("column #8 (DatExp)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value7.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value8.Length > 30)
                    {
                        MessageBox.Show("column #9 (AppSignedDate)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value8.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value9.Length > 30)
                    {
                        MessageBox.Show("column #10 (AppRcvDate)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value9.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value10.Length > 60)
                    {
                        MessageBox.Show("column #11 (Holder)  needs to be 60 or less characters.  At line " + (i + 1) + " you have a value that is " + value10.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value11.Length > 40)
                    {
                        MessageBox.Show("column #12 (HolderFirstName)  needs to be 40 or less characters.  At line " + (i + 1) + " you have a value that is " + value11.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value12.Length > 16)
                    {
                        MessageBox.Show("column #13 (HolderMiddleInitial)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value12.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value13.Length > 60)
                    {
                        MessageBox.Show("column #14 (HolderLastName)  needs to be 60 or less characters.  At line " + (i + 1) + " you have a value that is " + value13.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value14.Length > 60)
                    {
                        MessageBox.Show("column #15 (HolderStreet)  needs to be 60 or less characters.  At line " + (i + 1) + " you have a value that is " + value14.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value15.Length > 30)
                    {
                        MessageBox.Show("column #16 (HolderStreet2)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value15.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value16.Length > 40)
                    {
                        MessageBox.Show("column #17 (HolderCity)  needs to be 40 or less characters.  At line " + (i + 1) + " you have a value that is " + value16.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value17.Length > 6)
                    {
                        MessageBox.Show("column #18 (HolderState)  needs to be 6 or less characters.  At line " + (i + 1) + " you have a value that is " + value17.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value18.Length > 16)
                    {
                        MessageBox.Show("column #19 (HolderZip)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value18.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value19.Length > 40)
                    {
                        MessageBox.Show("column #20 (CountyCode)  needs to be 40 or less characters.  At line " + (i + 1) + " you have a value that is " + value19.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value20.Length > 20)
                    {
                        MessageBox.Show("column #21 (HolderPhone)  needs to be 20 or less characters.  At line " + (i + 1) + " you have a value that is " + value20.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value21.Length > 30)
                    {
                        MessageBox.Show("column #22 (HolderDOB)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value21.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value22.Length > 20)
                    {
                        MessageBox.Show("column #23 (HolderSSN)  needs to be 20 or less characters.  At line " + (i + 1) + " you have a value that is " + value22.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value23.Length > 30)
                    {
                        MessageBox.Show("column #24 (PolState)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value23.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value24.Length > 8)
                    {
                        MessageBox.Show("column #25 (DualCoverage)  needs to be 8 or less characters.  At line " + (i + 1) + " you have a value that is " + value24.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value25.Length > 16)
                    {
                        MessageBox.Show("column #26 (BrokerId)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value25.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value26.Length > 60)
                    {
                        MessageBox.Show("column #27 (TermType)  needs to be 60 or less characters.  At line " + (i + 1) + " you have a value that is " + value26.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value27.Length > 16)
                    {
                        MessageBox.Show("column #28 (ProCode)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value27.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value28.Length > 16)
                    {
                        MessageBox.Show("column #29 (BrokerId2)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value28.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value29.Length > 3.2)
                    {
                        MessageBox.Show("column #30 (PrimaryBrokerPct)  needs to be 3.2 or less characters.  At line " + (i + 1) + " you have a value that is " + value29.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value30.Length > 3.2)
                    {
                        MessageBox.Show("column #31 (SecondaryBrokerPct)  needs to be 3.2 or less characters.  At line " + (i + 1) + " you have a value that is " + value30.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value31.Length > 16)
                    {
                        MessageBox.Show("column #32 (ReferralId)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value31.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value32.Length > 5)
                    {
                        MessageBox.Show("column #33 (BusType)  needs to be 5 or less characters.  At line " + (i + 1) + " you have a value that is " + value32.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value33.Length > 30)
                    {
                        MessageBox.Show("column #34 (GroupId)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value33.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value34.Length > 40)
                    {
                        MessageBox.Show("column #35 (CustomerRegion)  needs to be 40 or less characters.  At line " + (i + 1) + " you have a value that is " + value34.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value35.Length > 20)
                    {
                        MessageBox.Show("column #36 (AppSource)  needs to be 20 or less characters.  At line " + (i + 1) + " you have a value that is " + value35.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    if (value36.Length > 30)
                    {
                        MessageBox.Show("column #37 (HolderDOD)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value36.Length + " characters long.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Medicare file is OK", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    return;
                }
            }
        }
        private void medicareButtonCreateFile_Click(object sender, EventArgs e)
        {
            System.IO.Directory.CreateDirectory("C:\\Program Files (x86)\\CCDataTool\\Medicare Error Files");
            string path = @"C:\\Program Files (x86)\\CCDataTool\\Medicare Error Files\\CCDataTool_MEF_" + DateTime.Now.ToString("MM_dd_yyyy_HHmmss") + ".txt";
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
            {
                using (TextWriter tw = new StreamWriter(fs))
                {

                    tw.WriteLine("CCDataTool - Beginning of Medicare Error File");
                    tw.WriteLine("Reading file...");
                    tw.WriteLine(".");
                    tw.WriteLine(".");
                    tw.WriteLine(".");
                    tw.WriteLine(".");

                    if (dataGridView1.ColumnCount != 37)
                    {
                        tw.WriteLine("Medicare files need 37 columns. You have " + dataGridView1.ColumnCount + ".");
                    }
                    //column 1 -required
                    try
                    {

                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value0 = dataGridView1.Rows[i].Cells[0].Value.ToString();

                            if (string.IsNullOrWhiteSpace(value0))
                            {
                                tw.WriteLine("NULL value found in column #1 (CustomerId)  at line " + (i + 1) + ". This is a required field.");
                            }
                            if (value0.Length > 30)
                            {
                                tw.WriteLine("column #1 (CustomerId)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #1 check...done."); }
                    //column 2 -required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value1 = dataGridView1.Rows[i].Cells[1].Value.ToString();

                            if (string.IsNullOrWhiteSpace(value1))
                            {
                                tw.WriteLine("NULL value found in column #2 (ContractNbr)  at line " + (i + 1) + ". This is a required field.");
                            }
                            if (value1.Length > 10)
                            {
                                tw.WriteLine("column #2 (ContractNbr)  needs to be 10 or less characters.  At line " + (i + 1) + " you have a value that is " + value1.Length + " characters long.");
                                return;
                            }
                        }
                    }
                    catch { tw.WriteLine("column #2 check...done."); }
                    //column 3 -required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value2 = dataGridView1.Rows[i].Cells[2].Value.ToString();
                            if (string.IsNullOrWhiteSpace(value2))
                            {
                                tw.WriteLine("NULL value found in column #3 (PBP)  at line " + (i + 1) + ". This is a required field.");
                            }
                            if (value2.Length > 10)
                            {
                                tw.WriteLine("column #3 (PBP)  needs to be 10 or less characters.  At line " + (i + 1) + " you have a value that is " + value2.Length + " characters long.");
                                return;
                            }
                        }
                    }
                    catch { tw.WriteLine("column #3 check...done."); }
                    //column 4 -required
                    try
                    {

                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value3 = dataGridView1.Rows[i].Cells[3].Value.ToString();
                            if (string.IsNullOrWhiteSpace(value3))
                            {
                                tw.WriteLine("NULL value found in column #4 (HICN)  at line " + (i + 1) + ". This is a required field.");
                            }
                            if (value3.Length > 20)
                            {
                                tw.WriteLine("column #4 (HICN)  needs to be 20 or less characters.  At line " + (i + 1) + " you have a value that is " + value3.Length + " characters long.");
                                return;
                            }

                        }
                    }
                    catch { tw.WriteLine("column #4 check...done."); }
                    //column 5 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value4 = dataGridView1.Rows[i].Cells[4].Value.ToString();

                            if (value4.Length > 20)
                            {
                                tw.WriteLine("column #5 (OED)  int");
                                return;
                            }
                        }
                    }
                    catch { tw.WriteLine("column #5 check...done."); }
                    //column 6 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value5 = dataGridView1.Rows[i].Cells[5].Value.ToString();

                            if (value5.Length > 20)
                            {
                                tw.WriteLine("column #6 (CMSOED)  int");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #6 check...done."); }
                    //column 7 -required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value6 = dataGridView1.Rows[i].Cells[6].Value.ToString();

                            if (string.IsNullOrWhiteSpace(value6))
                            {

                                tw.WriteLine("NULL value found in column #7 (DatEff)  at line " + (i + 1) + ". This is a required field.");
                            }
                            if (value6.Length > 20)
                            {
                                tw.WriteLine("column #7 (DatEff)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value6.Length + " characters long.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #7 check...done."); }
                    //column 8 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value7 = dataGridView1.Rows[i].Cells[7].Value.ToString();

                            if (value7.Length > 20)
                            {
                                tw.WriteLine("column #8 (DatExp)  int");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #8 check...done."); }
                    //column 9 -required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value8 = dataGridView1.Rows[i].Cells[8].Value.ToString();

                            if (string.IsNullOrWhiteSpace(value8))
                            {
                                tw.WriteLine("NULL value found in column #9 (AppSignedDate)  at line " + (i + 1) + ". This is a required field.");

                            }
                            if (value8.Length > 20)
                            {
                                tw.WriteLine("column #9 (AppSignedDate)  int-length for this?");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #9 check...done."); }
                    //column 10 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value9 = dataGridView1.Rows[i].Cells[9].Value.ToString();

                            if (value9.Length > 20)
                            {
                                tw.WriteLine("column #10 (AppRcvDate)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value9.Length + " characters long.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #10 check...done."); }
                    //column 11 -required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value10 = dataGridView1.Rows[i].Cells[10].Value.ToString();

                            if (string.IsNullOrWhiteSpace(value10))
                            {
                                tw.WriteLine("NULL value found in column #11 (Holder)  at line " + (i + 1) + ". This is a required field.");
                            }
                            if (value10.Length > 60)
                            {
                                tw.WriteLine("column #11 (Holder)  needs to be 60 or less characters.  At line " + (i + 1) + " you have a value that is " + value10.Length + " characters long.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #11 check...done."); }
                    //column 12 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value11 = dataGridView1.Rows[i].Cells[11].Value.ToString();

                            if (value11.Length > 60)
                            {
                                tw.WriteLine("column #12 (HolderFirstName)  needs to be 40 or less characters.  At line " + (i + 1) + " you have a value that is " + value11.Length + " characters long.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #12 check...done."); }
                    //column 13 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value12 = dataGridView1.Rows[i].Cells[12].Value.ToString();

                            if (value12.Length > 60)
                            {
                                tw.WriteLine("column #13 (HolderMiddleInitial)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value12.Length + " characters long.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #13 check...done."); }
                    //column 14 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value13 = dataGridView1.Rows[i].Cells[13].Value.ToString();

                            if (value13.Length > 60)
                            {
                                tw.WriteLine("column #14 (HolderLastName)  needs to be 60 or less characters.  At line " + (i + 1) + " you have a value that is " + value13.Length + " characters long.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #14 check...done."); }
                    //column 15 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value14 = dataGridView1.Rows[i].Cells[14].Value.ToString();

                            if (value14.Length > 60)
                            {
                                tw.WriteLine("column #15 (HolderStreet)  needs to be 60 or less characters.  At line " + (i + 1) + " you have a value that is " + value14.Length + " characters long.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #15 check...done."); }
                    //column 16 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value15 = dataGridView1.Rows[i].Cells[15].Value.ToString();

                            if (value15.Length > 60)
                            {
                                tw.WriteLine("column #16 (HolderStreet2)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value15.Length + " characters long.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #16 check...done."); }
                    //column 17 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value16 = dataGridView1.Rows[i].Cells[16].Value.ToString();

                            if (value16.Length > 60)
                            {
                                tw.WriteLine("column #17 (HolderCity)  needs to be 40 or less characters.  At line " + (i + 1) + " you have a value that is " + value16.Length + " characters long.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #17 check...done."); }
                    //column 18 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value17 = dataGridView1.Rows[i].Cells[17].Value.ToString();

                            if (value17.Length > 60)
                            {
                                tw.WriteLine("column #18 (HolderState)  needs to be 6 or less characters.  At line " + (i + 1) + " you have a value that is " + value17.Length + " characters long.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #18 check...done."); }
                    //column 19 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value18 = dataGridView1.Rows[i].Cells[18].Value.ToString();

                            if (value18.Length > 60)
                            {
                                tw.WriteLine("column #19 (HolderZip)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value18.Length + " characters long.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #19 check...done."); }
                    //column 20 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value19 = dataGridView1.Rows[i].Cells[19].Value.ToString();

                            if (value19.Length > 60)
                            {
                                tw.WriteLine("column #20 (CountyCode)  needs to be 40 or less characters.  At line " + (i + 1) + " you have a value that is " + value19.Length + " characters long.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #20 check...done."); }
                    //column 21 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value20 = dataGridView1.Rows[i].Cells[20].Value.ToString();

                            if (value20.Length > 60)
                            {
                                tw.WriteLine("column #21 (HolderPhone)  needs to be 20 or less characters.  At line " + (i + 1) + " you have a value that is " + value20.Length + " characters long.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #21 check...done."); }
                    //column 22 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value21 = dataGridView1.Rows[i].Cells[21].Value.ToString();

                            if (value21.Length > 60)
                            {
                                tw.WriteLine("column #22 (HolderDOB)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value21.Length + " characters long.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #22 check...done."); }
                    //column 23 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value22 = dataGridView1.Rows[i].Cells[22].Value.ToString();

                            if (value22.Length > 60)
                            {
                                tw.WriteLine("column #23 (HolderSSN)  needs to be 20 or less characters.  At line " + (i + 1) + " you have a value that is " + value22.Length + " characters long.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #23 check...done."); }
                    //column 24 -required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value23 = dataGridView1.Rows[i].Cells[23].Value.ToString();

                            if (string.IsNullOrWhiteSpace(value23))
                            {
                                tw.WriteLine("NULL value found in column #24 (PolState)  at line " + (i + 1) + ". This is a required field.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #24 check...done."); }
                    //column 25 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value24 = dataGridView1.Rows[i].Cells[24].Value.ToString();

                            if (value24.Length > 60)
                            {
                                tw.WriteLine("NULL value found in column #24 (PolState)  at line " + (i + 1) + ". This is a required field.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #25 check...done."); }
                    //column 26 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value25 = dataGridView1.Rows[i].Cells[25].Value.ToString();

                            if (value25.Length > 60)
                            {
                                tw.WriteLine("NULL value found in column #24 (PolState)  at line " + (i + 1) + ". This is a required field.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #26 check...done."); }
                    //column 27 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value26 = dataGridView1.Rows[i].Cells[26].Value.ToString();

                            if (value26.Length > 60)
                            {
                                tw.WriteLine("NULL value found in column #24 (PolState)  at line " + (i + 1) + ". This is a required field.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #27 check...done."); }
                    //column 28 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value27 = dataGridView1.Rows[i].Cells[27].Value.ToString();

                            if (value27.Length > 60)
                            {
                                tw.WriteLine("NULL value found in column #24 (PolState)  at line " + (i + 1) + ". This is a required field.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #28 check...done."); }
                    //column 29 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value28 = dataGridView1.Rows[i].Cells[28].Value.ToString();

                            if (value28.Length > 60)
                            {
                                tw.WriteLine("NULL value found in column #24 (PolState)  at line " + (i + 1) + ". This is a required field.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #29 check...done."); }
                    //column 30 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value29 = dataGridView1.Rows[i].Cells[29].Value.ToString();

                            if (value29.Length > 60)
                            {
                                tw.WriteLine("NULL value found in column #24 (PolState)  at line " + (i + 1) + ". This is a required field.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #30 check...done."); }
                    //column 31 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value30 = dataGridView1.Rows[i].Cells[30].Value.ToString();

                            if (value30.Length > 60)
                            {
                                tw.WriteLine("NULL value found in column #24 (PolState)  at line " + (i + 1) + ". This is a required field.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #31 check...done."); }
                    //column 32 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value31 = dataGridView1.Rows[i].Cells[31].Value.ToString();

                            if (value31.Length > 60)
                            {
                                tw.WriteLine("NULL value found in column #24 (PolState)  at line " + (i + 1) + ". This is a required field.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #32 check...done."); }
                    //column 33 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value32 = dataGridView1.Rows[i].Cells[32].Value.ToString();

                            if (value32.Length > 60)
                            {
                                tw.WriteLine("NULL value found in column #24 (PolState)  at line " + (i + 1) + ". This is a required field.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #33 check...done."); }
                    //column 34 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value33 = dataGridView1.Rows[i].Cells[33].Value.ToString();

                            if (value33.Length > 60)
                            {
                                tw.WriteLine("NULL value found in column #24 (PolState)  at line " + (i + 1) + ". This is a required field.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #34 check...done."); }
                    //column 35 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value34 = dataGridView1.Rows[i].Cells[34].Value.ToString();

                            if (value34.Length > 60)
                            {
                                tw.WriteLine("NULL value found in column #24 (PolState)  at line " + (i + 1) + ". This is a required field.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #35 check...done."); }
                    //column 36 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value35 = dataGridView1.Rows[i].Cells[35].Value.ToString();

                            if (value35.Length > 60)
                            {
                                tw.WriteLine("NULL value found in column #24 (PolState)  at line " + (i + 1) + ". This is a required field.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #36 check...done."); }
                    //column 37 -not required
                    try
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            var value36 = dataGridView1.Rows[i].Cells[36].Value.ToString();

                            if (value36.Length > 60)
                            {
                                tw.WriteLine("NULL value found in column #24 (PolState)  at line " + (i + 1) + ". This is a required field.");
                            }
                        }
                    }
                    catch { tw.WriteLine("column #37 check...done."); }
                    tw.WriteLine("EOF.");
                }


            }
            MessageBox.Show("Medicare error file has been created. \nLocation: C:\\Program Files (x86)\\CCDataTool\\Medicare Error Files", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
        }



        //------------------MEDICARE CHECKER END------------------------------------------------------

    }
}