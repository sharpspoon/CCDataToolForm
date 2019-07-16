using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Net;
using System.Text;
using PgpCore;
using System.Collections;
using System.Collections.Generic;

namespace SAPDataAnalysisTool
{

    public partial class SAPDataAnalysisTool : Form
    {
        /*
         * ############################################################################################   
         * ############################################################################################
         * ####################PRODUCTION CODE BEGIN###################################################
         * ############################################################################################
         * ############################################################################################
        */

        //*********************************************************************************************
        //*********************************GLOBAL******************************************************
        //*********************************************************************************************

        //------------------BUTTON MOUSE LOGIC START------------------------------------------------------

        private void moveUpPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.moveUpPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_move_up3));
        }

        private void moveUpPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.moveUpPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_move_up2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.moveUpPictureBox, "Run the tool!");
        }

        private void moveUpPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.moveUpPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_move_up));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.moveUpPictureBox, "Run the tool!");
        }

        private void moveUpPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.moveUpPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_move_up));
        }

        private void moveDownPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.moveDownPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_move_down3));
        }

        private void moveDownPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.moveDownPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_move_down2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.moveDownPictureBox, "Run the tool!");
        }

        private void moveDownPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.moveDownPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_move_down));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.moveDownPictureBox, "Run the tool!");
        }

        private void moveDownPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.moveDownPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_move_down));
        }

        private void goButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.goButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.goButtonPictureBox, "Run the tool!");
        }

        private void goButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.goButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.goButtonPictureBox, "Run the tool!");
        }

        private void goButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.goButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go3));
        }

        private void goButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.goButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
        }

        private void csvButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.csvButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_csv3));
        }

        private void csvButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.csvButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_csv2));
        }

        private void csvButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.csvButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_csv));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.csvButtonPictureBox, "Open a CSV file.");
        }

        private void csvButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.csvButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_csv));
        }

        private void xmlButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.xmlButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xml3));
        }

        private void xmlButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.xmlButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xml2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.xmlButtonPictureBox, "Open an XML file.");
        }

        private void xmlButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.xmlButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xml));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.xmlButtonPictureBox, "Open an XML file.");
        }

        private void xmlButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.xmlButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xml));
        }

        private void txtCommaButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.txtCommaButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_txt_comma3));
        }

        private void txtCommaButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.txtCommaButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_txt_comma2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.txtCommaButtonPictureBox, "Open a Text Comma file.");
        }

        private void txtCommaButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.txtCommaButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_txt_comma));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.txtCommaButtonPictureBox, "Open a Text Comma file.");
        }

        private void txtCommaButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.txtCommaButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_txt_comma));
        }

        private void xlsButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.xlsButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xls3));
        }

        private void xlsButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.xlsButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xls2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.xlsButtonPictureBox, "Open an XLS file.");
        }

        private void xlsButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.xlsButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xls));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.xlsButtonPictureBox, "Open an XLS file.");
        }

        private void xlsButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.xlsButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xls));
        }

        private void txtPipePictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.txtPipePictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_txt_pipe3));
        }

        private void txtPipePictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.txtPipePictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_txt_pipe2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.txtPipePictureBox, "Open a Text Pipe file.");
        }

        private void txtPipePictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.txtPipePictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_txt_pipe));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.txtPipePictureBox, "Open a Text Pipe file.");
        }

        private void txtPipePictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.txtPipePictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_txt_pipe));
        }

        private void clearResultsPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.clearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results3));
        }

        private void clearResultsPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.clearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearResultsPictureBox, "Clear the results.");
        }

        private void clearResultsPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.clearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearResultsPictureBox, "Clear the results.");
        }

        private void clearResultsPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.clearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results));
        }

        private void exportResultsPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.exportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results3));
        }

        private void exportResultsPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.exportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.exportResultsPictureBox, "Export the results.");
        }

        private void exportResultsPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.exportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.exportResultsPictureBox, "Export the results.");
        }

        private void exportResultsPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.exportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results));
        }

        private void envChangesGoPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.envChangesGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go3));
        }

        private void envChangesGoPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.envChangesGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.envChangesGoPictureBox, "Run the tool!");
        }

        private void envChangesGoPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.envChangesGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.envChangesGoPictureBox, "Run the tool!");
        }

        private void envChangesGoPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.envChangesGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
        }

        private void clearResultsPictureBox_Click(object sender, EventArgs e)
        {
            envChangesRichTextBox.Clear();
        }

        private void apiGoPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.apiGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go3));
        }

        private void apiGoPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.apiGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.apiGoPictureBox, "Run the tool!");
        }

        private void apiGoPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.apiGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.apiGoPictureBox, "Run the tool!");
        }

        private void apiGoPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.apiGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
        }

        private void apiExportResultsPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.apiExportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results3));
        }

        private void apiExportResultsPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.apiExportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.apiExportResultsPictureBox, "Export the results.");
        }

        private void apiExportResultsPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.apiExportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.apiExportResultsPictureBox, "Export the results.");
        }

        private void apiExportResultsPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.apiExportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results));
        }

        private void apiClearResultsPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.apiClearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results3));
        }

        private void apiClearResultsPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.apiClearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.apiClearResultsPictureBox, "Clear the results.");
        }

        private void apiClearResultsPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.apiClearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.apiClearResultsPictureBox, "Clear the results.");
        }

        private void apiClearResultsPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.apiClearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results));
        }

        private void benchmarkExportResultsPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.benchmarkExportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results3));
        }

        private void benchmarkExportResultsPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.benchmarkExportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.benchmarkExportResultsPictureBox, "Export the results.");
        }

        private void benchmarkExportResultsPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.benchmarkExportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.benchmarkExportResultsPictureBox, "Export the results.");
        }

        private void benchmarkExportResultsPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.benchmarkExportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results));
        }

        private void benchmarkClearResultsPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.benchmarkClearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results3));
        }

        private void benchmarkClearResultsPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.benchmarkClearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.benchmarkClearResultsPictureBox, "Clear the results.");
        }

        private void benchmarkClearResultsPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.benchmarkClearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.benchmarkClearResultsPictureBox, "Clear the results.");
        }

        private void benchmarkClearResultsPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.benchmarkClearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results));
        }

        private void cellLengthCheckerGoButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.cellLengthCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go3));
        }

        private void cellLengthCheckerGoButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.cellLengthCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.cellLengthCheckerGoButtonPictureBox, "Run the check.");
        }

        private void cellLengthCheckerGoButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.cellLengthCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.cellLengthCheckerGoButtonPictureBox, "Run the check.");
        }

        private void cellLengthCheckerGoButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.cellLengthCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go));
        }

        private void clearAllCellLengthCheckerPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.clearAllCellLengthCheckerPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear3));
        }

        private void clearAllCellLengthCheckerPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.clearAllCellLengthCheckerPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearAllCellLengthCheckerPictureBox, "Clear all.");
        }

        private void clearAllCellLengthCheckerPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.clearAllCellLengthCheckerPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearAllCellLengthCheckerPictureBox, "Clear all.");
        }

        private void clearAllCellLengthCheckerPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.clearAllCellLengthCheckerPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear));
        }

        private void selectAllCellLengthCheckerPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.selectAllCellLengthCheckerPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all3));
        }

        private void selectAllCellLengthCheckerPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.selectAllCellLengthCheckerPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.selectAllCellLengthCheckerPictureBox, "Select all.");
        }

        private void selectAllCellLengthCheckerPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.selectAllCellLengthCheckerPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.selectAllCellLengthCheckerPictureBox, "Select all.");
        }

        private void selectAllCellLengthCheckerPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.selectAllCellLengthCheckerPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all));
        }

        private void nullCheckerGoButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.nullCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go3));
        }

        private void nullCheckerGoButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.nullCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.nullCheckerGoButtonPictureBox, "Run the check.");
        }

        private void nullCheckerGoButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.nullCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.nullCheckerGoButtonPictureBox, "Run the check.");
        }

        private void nullCheckerGoButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.nullCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go));
        }

        private void clearAllNullCheckerButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.clearAllNullCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear3));
        }

        private void clearAllNullCheckerButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.clearAllNullCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearAllNullCheckerButtonPictureBox, "Clear all.");
        }

        private void clearAllNullCheckerButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.clearAllNullCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearAllNullCheckerButtonPictureBox, "Clear all.");
        }

        private void clearAllNullCheckerButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.clearAllNullCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear));
        }

        private void selectAllNullCheckerButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.selectAllNullCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all3));
        }

        private void selectAllNullCheckerButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.selectAllNullCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.selectAllNullCheckerButtonPictureBox, "Select all.");
        }

        private void selectAllNullCheckerButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.selectAllNullCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.selectAllNullCheckerButtonPictureBox, "Select all.");
        }

        private void selectAllNullCheckerButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.selectAllNullCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all));
        }

        private void specialCharacterCheckerGoButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.specialCharacterCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go3));
        }

        private void specialCharacterCheckerGoButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.specialCharacterCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.specialCharacterCheckerGoButtonPictureBox, "Run the check.");
        }

        private void specialCharacterCheckerGoButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.specialCharacterCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.specialCharacterCheckerGoButtonPictureBox, "Run the check.");
        }

        private void specialCharacterCheckerGoButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.specialCharacterCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go));
        }

        private void clearAllSpecialCharacterCheckerButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.clearAllSpecialCharacterCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear3));
        }

        private void clearAllSpecialCharacterCheckerButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.clearAllSpecialCharacterCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearAllSpecialCharacterCheckerButtonPictureBox, "Clear all.");
        }

        private void clearAllSpecialCharacterCheckerButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.clearAllSpecialCharacterCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearAllSpecialCharacterCheckerButtonPictureBox, "Clear all.");
        }

        private void clearAllSpecialCharacterCheckerButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.clearAllSpecialCharacterCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear));
        }

        private void selectAllSpecialCharacterCheckerButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.selectAllSpecialCharacterCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all3));
        }

        private void selectAllSpecialCharacterCheckerButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.selectAllSpecialCharacterCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.selectAllSpecialCharacterCheckerButtonPictureBox, "Select all.");
        }

        private void selectAllSpecialCharacterCheckerButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.selectAllSpecialCharacterCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.selectAllSpecialCharacterCheckerButtonPictureBox, "Select all.");
        }

        private void selectAllSpecialCharacterCheckerButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.selectAllSpecialCharacterCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all));
        }

        private void dateCheckerGoButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.dateCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go3));
        }

        private void dateCheckerGoButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.dateCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.dateCheckerGoButtonPictureBox, "Run the check.");
        }

        private void dateCheckerGoButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.dateCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.dateCheckerGoButtonPictureBox, "Run the check.");
        }

        private void dateCheckerGoButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.dateCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go));
        }

        private void clearAllDateCheckerButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.clearAllDateCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear3));
        }

        private void clearAllDateCheckerButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.clearAllDateCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearAllDateCheckerButtonPictureBox, "Clear all.");
        }

        private void clearAllDateCheckerButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.clearAllDateCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearAllDateCheckerButtonPictureBox, "Clear all.");
        }

        private void clearAllDateCheckerButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.clearAllDateCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear));
        }

        private void selectAllDateCheckerButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.selectAllDateCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all3));
        }

        private void selectAllDateCheckerButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.selectAllDateCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.selectAllDateCheckerButtonPictureBox, "Select all.");
        }

        private void selectAllDateCheckerButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.selectAllDateCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.selectAllDateCheckerButtonPictureBox, "Select all.");
        }

        private void selectAllDateCheckerButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.selectAllDateCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all));
        }

        private void fileSweepUploadFilesPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.fileSweepUploadFilesPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_upload_files3));
        }

        private void fileSweepUploadFilesPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.fileSweepUploadFilesPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_upload_files2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.fileSweepUploadFilesPictureBox, "Upload file(s).");
        }

        private void fileSweepUploadFilesPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.fileSweepUploadFilesPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_upload_files));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.fileSweepUploadFilesPictureBox, "Upload file(s).");
        }

        private void fileSweepUploadFilesPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.fileSweepUploadFilesPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_upload_files));
        }

        private void fileSweepGoPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.fileSweepGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go3));
        }

        private void fileSweepGoPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.fileSweepGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.fileSweepGoPictureBox, "Run the tool!");
        }

        private void fileSweepGoPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.fileSweepGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.fileSweepGoPictureBox, "Run the tool!");
        }

        private void fileSweepGoPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.fileSweepGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
        }

        //------------------BUTTON MOUSE LOGIC END------------------------------------------------------

        public SAPDataAnalysisTool()
        {
            InitializeComponent();
            dateComboBox1.SelectedIndex = 12;
            dateComboBox2.SelectedIndex = 5;
            dateComboBox3.SelectedIndex = 1;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            unableToRegUserToolStripStatusLabel.Text = @"TALLYCENTRAL\" + Environment.UserName;
        }

        Loading loading = new Loading();

        //------------------FORM DRAG LOGIC START------------------------------------------------------
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();
        private void Form1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        //------------------FORM DRAG LOGIC END------------------------------------------------------

        //------------------CROW NUMBER LOGIC START------------------------------------------------------
        private void dgvUserDetails_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(importedfileDataGridView.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }
        //------------------CROW NUMBER LOGIC END------------------------------------------------------

        private void Form_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        //------------------TOOLTIP LOGIC START------------------------------------------------------

        ToolTip tt = new ToolTip();

        private void serverSelect_MouseHover(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip1 = new System.Windows.Forms.ToolTip();
            ToolTip1.SetToolTip(this.serverSelect, "Select your ICM server.");
        }

        private void databaseSelect_MouseHover(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.databaseSelect, "Select your ICM database.");
        }

        private void ifSelect_MouseHover(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.ifSelect, "Select your Import Format.");

        }

        private void groupBox7_MouseHover(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.importFormatServerSelectGroupBox, "Select your Server/Database/Import Format.");
        }

        private void reqListBox_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.reqListBox, "Select your required Import Format fields.");
        }

        private void groupBox1_MouseHover(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.importFormatSelectRequiredFieldsGroupBox, "Select your required Import Format fields.");
        }

        private void dateListBox_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.dateListBox, "Select the columns your created date format should apply to.");
        }

        private void dateComboBox1_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.dateComboBox1, "Use this dropdown to build your date format.");
        }

        private void dateComboBox2_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.dateComboBox2, "Use this dropdown to build your date format.");
        }

        private void dateComboBox3_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.dateComboBox3, "Use this dropdown to build your date format.");
        }

        private void dateComboBoxSeperator_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.dateComboBoxSeperator, "Do you want to use a seperator?");
        }

        private void dateFormat_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.dateFormat, "This is the current date format you built");
        }

        private void checkBox2_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.importFormatFindNullCheckbox, "Do you want to find NULLs in the date column?");
        }

        private void tableSelect_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.tableSelect, "Use this dropdown to check any table within your selected database.");
        }

        //------------------TOOLTIP LOGIC END------------------------------------------------------

        private void toolStripStatusLabel18_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.sap.com/index.html");
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
    (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        //*********************************************************************************************
        //*********************************/GLOBAL*****************************************************
        //*********************************************************************************************

        //*********************************************************************************************
        //*********************************HEADER MENU*************************************************
        //*********************************************************************************************

        //------------------SAP LOG OPEN START------------------------------------------------------
        private void cCDataToolLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            try
            {
                Process.Start(Application.UserAppDataPath + @"\Logs");
                progressBar1.MarqueeAnimationSpeed = 0;
            }
            catch
            {
                progressBar1.MarqueeAnimationSpeed = 0;
            }
        }
        //------------------SAP LOG OPEN END------------------------------------------------------

        //------------------OPEN/SAVE XML START------------------------------------------------------
        private void menu_Open_Xml_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            try
            {
                DataSet dataSet = new DataSet();
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "XML | *.xml", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        dataSet.ReadXml(ofd.FileName);
                        importedfileDataGridView.DataSource = dataSet.Tables[0];

                        importFormatActualFileNameToolStripStatusLabel.Text = ofd.FileName;
                        ifRowCounterToolStripStatusLabel.Text = importedfileDataGridView[0, importedfileDataGridView.Rows.Count - 1].Value.ToString();
                        systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading XML: " + ofd.FileName + "...Done.");
                        ifRowCountLabelToolStripStatusLabel.Visible = true;
                        ifRowCounterToolStripStatusLabel.Visible = true;
                        seperator3ToolStripStatusLabel.Visible = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.MarqueeAnimationSpeed = 0;
        }
        //------------------OPEN/SAVE XML END------------------------------------------------------

        //------------------OPEN/SAVE XLS START------------------------------------------------------

        private void menu_Open_Xls_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            try
            {
                OpenFileDialog openfile1 = new OpenFileDialog();
                if (openfile1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    this.importFormatActualFileNameToolStripStatusLabel.Text = openfile1.FileName;
                }
                {
                    string pathconn = "Provider = Microsoft.jet.OLEDB.4.0; Data source=" + importFormatActualFileNameToolStripStatusLabel.Text + ";Extended Properties=\"Excel 8.0;HDR= yes;\";";
                    OleDbConnection conn = new OleDbConnection(pathconn);
                    OleDbDataAdapter MyDataAdapter = new OleDbDataAdapter("Select * from [Sheet1$]", conn);
                    DataTable dt = new DataTable();
                    MyDataAdapter.Fill(dt);
                    importedfileDataGridView.DataSource = dt;
                }
            }
            catch { }
            progressBar1.MarqueeAnimationSpeed = 0;
        }
        //------------------OPEN/SAVE XLS END------------------------------------------------------

        //------------------CUT, COPY, PASTE START------------------------------------------------------
        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Control ctrl = this.ActiveControl;
            if (ctrl != null)
            {
                if (ctrl is TextBox)
                {
                    TextBox tx = (TextBox)ctrl;
                    tx.Copy();
                }
            }
        }
        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Control ctrl = this.ActiveControl;
            if (ctrl != null)
            {
                if (ctrl is TextBox)
                {
                    TextBox tx = (TextBox)ctrl;
                    tx.Cut();
                }
            }
        }
        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Control ctrl = this.ActiveControl;
            if (ctrl != null)
            {
                if (ctrl is TextBox)
                {
                    TextBox tx = (TextBox)ctrl;
                    tx.Paste();
                }
            }
        }
        //------------------CUT, COPY, PASTE END------------------------------------------------------

        //------------------TOOLSTRIP MINIMIZE, MAXIMIZE, CLOSE START------------------------------------------------------
        private void toolStripMenuItemClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void toolStripMenuItemMaximize_Click(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized)
            {
                this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;
                this.WindowState = FormWindowState.Normal;
            }
            else
            {
                this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;
                this.WindowState = FormWindowState.Maximized;
            }
        }
        private void toolStripMenuItemMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        //------------------TOOLSTRIP MINIMIZE, MAXIMIZE, CLOSE END------------------------------------------------------

        //------------------PRINT DOCUMENT START------------------------------------------------------
        Bitmap bitmap;
        private void btnPrint_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            if (importedfileDataGridView.Rows.Count == 0 || importedfileDataGridView.Rows == null)
            {
                MessageBox.Show("No data to print", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                //Resize DataGridView to full height.
                int height = importedfileDataGridView.Height;
                importedfileDataGridView.Height = importedfileDataGridView.RowCount * importedfileDataGridView.RowTemplate.Height;

                //Create a Bitmap and draw the DataGridView on it.
                bitmap = new Bitmap(this.importedfileDataGridView.Width, this.importedfileDataGridView.Height);
                importedfileDataGridView.DrawToBitmap(bitmap, new Rectangle(0, 0, this.importedfileDataGridView.Width, this.importedfileDataGridView.Height));

                //Resize DataGridView back to original height.
                importedfileDataGridView.Height = height;

                //Show the Print Preview Dialog.
                printPreviewDialog1.Document = printDocument1;
                printPreviewDialog1.PrintPreviewControl.Zoom = 1;
                printPreviewDialog1.ShowDialog();
            }
            progressBar1.MarqueeAnimationSpeed = 0;
        }
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            //Print the contents.
            e.Graphics.DrawImage(bitmap, 0, 0);
        }
        //------------------PRINT DOCUMENT END------------------------------------------------------

        //------------------OPEN/SAVE CSV START------------------------------------------------------
        private void menu_Open_Csv_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;

            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "CSV | *.csv", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        importedfileDataGridView.DataSource = ReadCsv(ofd.FileName);
                        importFormatActualFileNameToolStripStatusLabel.Text = ofd.FileName;
                        importFormatActualFileNameToolStripStatusLabel.Visible = true;
                        ifRowCounterToolStripStatusLabel.Text = importedfileDataGridView.Rows.Count.ToString();
                        ifRowCountLabelToolStripStatusLabel.Visible = true;
                        ifRowCounterToolStripStatusLabel.Visible = true;
                        seperator3ToolStripStatusLabel.Visible = true;
                        importFormatFileNameToolStripStatusLabel.Visible = true;
                        systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading CSV: " + ofd.FileName + "...Done.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            var importedFileArray = importedfileDataGridView.Columns.Cast<DataGridViewColumn>()
                .Select(x => x.HeaderCell.Value.ToString().Trim()).ToArray();
            dateCheckerListBox.Items.Clear();
            specialCharacterCheckerListBox.Items.Clear();
            nullCheckerListBox.Items.Clear();
            cellLengthCheckerListBox.Items.Clear();
            int a = 0;
            for (int i = 0; i < importedFileArray.Length; i++)
            {
                a++;

                specialCharacterCheckerListBox.Items.Add(a + ". " + importedFileArray[i].ToString());
                dateCheckerListBox.Items.Add(a + ". " + importedFileArray[i].ToString());
                nullCheckerListBox.Items.Add(a + ". " + importedFileArray[i].ToString());
                cellLengthCheckerListBox.Items.Add(a + ". " + importedFileArray[i].ToString());
            }
            progressBar1.MarqueeAnimationSpeed = 0;
        }
        public DataTable ReadCsv(string fileName)
        {
            importFormatProgressBar.Value = 0;
            importFormatProgressBar.Value = 20;
            System.Threading.Thread.Sleep(50);
            importFormatProgressBar.Value = 40;
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
            importFormatProgressBar.Value = 100;
            return dt;
        }
        //------------------OPEN/SAVE CSV END------------------------------------------------------

        //------------------ABOUT START------------------------------------------------------
        private void menu_About_Click(object sender, EventArgs e)
        {
            About about = new About();
            about.Show();
        }
        //------------------ABOUT END------------------------------------------------------

        //------------------EXIT APP ACTION START------------------------------------------------------
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            if (e.CloseReason == CloseReason.UserClosing)
            {
                DialogResult result = MessageBox.Show("Do you really want to exit?", "Data Analysis Tool", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    notifyIcon1.Visible = false;
                    notifyIcon1.Icon = null;
                    notifyIcon1.Dispose();
                    if (systemLogTextBox.Text == "")
                        Environment.Exit(0);
                    else
                    {
                        System.IO.Directory.CreateDirectory(Application.UserAppDataPath + @"\Logs");
                        string path = Application.UserAppDataPath + @"\Logs\DataAnalysisTool_Log_" + DateTime.Now.ToString("MM_dd_yyyy_HHmmss") + ".txt";
                        using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
                        {
                            using (TextWriter tw = new StreamWriter(fs))
                            {

                                tw.WriteLine("Data Analysis Tool - Activity Log");
                                tw.WriteLine("Log begin...");
                                tw.WriteLine(".");
                                tw.WriteLine(".");
                                tw.WriteLine(".");
                                tw.WriteLine(systemLogTextBox.Text);
                                tw.WriteLine("EOF.");
                            }
                        }
                        Environment.Exit(0);
                    }
                }
                else
                {
                    e.Cancel = true;
                }
            }
            else
            {
                e.Cancel = true;
            }
            progressBar1.MarqueeAnimationSpeed = 0;
        }
        //------------------EXIT APP ACTION END------------------------------------------------------

        //------------------ACKTEKSOFT LOGIN START------------------------------------------------------
        private void acteksoft_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 10;
            acteksoft actek = new acteksoft();
            while (Application.OpenForms.Count > 1)
            {
                Application.OpenForms[Application.OpenForms.Count - 1].Close();
            }
            actek.ShowDialog();
            progressBar1.MarqueeAnimationSpeed = 0;
        }
        //------------------ACKTEKSOFT LOGIN END------------------------------------------------------


        //*********************************************************************************************
        //*********************************/HEADER MENU************************************************
        //*********************************************************************************************

        //*********************************************************************************************
        //*********************************IMPORT FORMAT TAB*******************************************
        //*********************************************************************************************

        private void button7_Click(object sender, EventArgs e)
        {
            Process.Start(Application.UserAppDataPath + @"\IF_Error_Files");
        }

        private void tXTToolStripMenuItemComma_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "TXT | *.txt", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        importedfileDataGridView.DataSource = ReadTxtComma(ofd.FileName);
                        importFormatActualFileNameToolStripStatusLabel.Text = ofd.FileName;
                        importFormatActualFileNameToolStripStatusLabel.Visible = true;
                        ifRowCounterToolStripStatusLabel.Text = importedfileDataGridView.Rows.Count.ToString();
                        ifRowCountLabelToolStripStatusLabel.Visible = true;
                        ifRowCounterToolStripStatusLabel.Visible = true;
                        seperator3ToolStripStatusLabel.Visible = true;
                        importFormatFileNameToolStripStatusLabel.Visible = true;
                        systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading TXT: " + ofd.FileName + "...Done.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.MarqueeAnimationSpeed = 0;
        }

        //*********************************************************************************************
        //*********************************/IMPORT FORMAT TAB******************************************
        //*********************************************************************************************

        //*********************************************************************************************
        //*********************************CHECK TOOLS TAB*********************************************
        //*********************************************************************************************

        //------------------SELECT/CLEAR LIST BOX START------------------------------------------------------
        

        //------------------SELECT/CLEAR LIST BOX END------------------------------------------------------

        //------------------DATE CONVERTER START------------------------------------------------------

        //------------------DATE CONVERTER END------------------------------------------------------

        //------------------NULL CHECKER START------------------------------------------------------
        
        //------------------NULL CHECKER END------------------------------------------------------

        //------------------CELL LENGTH CHECKER START------------------------------------------------------
        

        //------------------CELL LENGTH CHECKER END------------------------------------------------------

        //------------------SPECIAL CHARACTER CHECKER START------------------------------------------------------
        

        //------------------SPECIAL CHARACTER CHECKER END------------------------------------------------------

        //*********************************************************************************************
        //*********************************/CHECK TOOLS TAB********************************************
        //*********************************************************************************************

        //*********************************************************************************************
        //*********************************SQL QUERY TAB**********************************************
        //*********************************************************************************************

        //*********************************************************************************************
        //*********************************/SQL QUERY TAB**********************************************
        //*********************************************************************************************

        //*********************************************************************************************
        //*********************************PAYOUT BENCHMARK TAB****************************************
        //*********************************************************************************************

        private void pendingRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (payoutTypeSelect.Text != "")
            {
                int value = payoutTypeSelect.SelectedIndex;
                payoutTypeSelect.SelectedIndex = -1;
                payoutTypeSelect.SelectedIndex = value;
            }
        }

        private void finalizedRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (payoutTypeSelect.Text != "")
            {

                int value = payoutTypeSelect.SelectedIndex;
                payoutTypeSelect.SelectedIndex = -1;
                payoutTypeSelect.SelectedIndex = value;
            }
        }

        private void reversedRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (payoutTypeSelect.Text != "")
            {

                int value = payoutTypeSelect.SelectedIndex;
                payoutTypeSelect.SelectedIndex = -1;
                payoutTypeSelect.SelectedIndex = value;
            }
        }

        private void payoutBenchmarkButton_Click(object sender, EventArgs e)
        {

            benchmarkProgressBar.Value = 0;
            benchmarkProgressBar.Value = 10;

            //global vars
            progressBar1.MarqueeAnimationSpeed = 1;
            if (serverSelect4.Text == "")

            {
                DialogResult result = MessageBox.Show("No server selected. \nPlease make sure you are connected to ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                benchmarkProgressBar.Value = 0;
                return;
            }

            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect4.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();

            //runlistnoroot
            var runListNoRoot = "";
            if (pendingRadioButton.Checked == true)
            {
                runListNoRoot = " USE " + databaseSelect4.Text + " select distinct rl.runlistnoroot from RunList rl left join rundet rd on rd.runlistno = rl.runlistno where rl.rectype='pay' and rd.ItemName = 'PayoutTypeNo' and rd.ItemValue = (select payouttypeno from PayoutType where payouttypeid = '" + payoutTypeSelect.Text + "') and rl.DatFrom = '" + payoutSelect.Text + "' and rl.finalizestatus='p'";
            }
            else if (finalizedRadioButton.Checked == true)
            {
                runListNoRoot = " USE " + databaseSelect4.Text + " select distinct rl.runlistnoroot from RunList rl left join rundet rd on rd.runlistno = rl.runlistno where rl.rectype='pay' and rd.ItemName = 'PayoutTypeNo' and rd.ItemValue = (select payouttypeno from PayoutType where payouttypeid = '" + payoutTypeSelect.Text + "') and rl.DatFrom = '" + payoutSelect.Text + "' and rl.finalizestatus='f'";
            }
            else if (reversedRadioButton.Checked == true)
            {
                runListNoRoot = " USE " + databaseSelect4.Text + " select distinct rl.runlistnoroot from RunList rl left join rundet rd on rd.runlistno = rl.runlistno where rl.rectype='pay' and rd.ItemName = 'PayoutTypeNo' and rd.ItemValue = (select payouttypeno from PayoutType where payouttypeid = '" + payoutTypeSelect.Text + "') and rl.DatFrom = '" + payoutSelect.Text + "' and rl.finalizestatus='r'";
            }
            var dataAdapter3 = new SqlDataAdapter(runListNoRoot, conn);
            var ds3 = new DataSet();
            dataAdapter3.Fill(ds3);
            stagedDataGridView.DataSource = ds3.Tables[0];
            var runListNo = stagedDataGridView.Rows[0].Cells[0].Value;

            //elapsed time
            var elapsedTime = " USE " + databaseSelect4.Text + " select distinct (elapsedtime / 1000) / 60 as name from RunList  where RunListNo = " + runListNo;
            var dataAdapter4 = new SqlDataAdapter(elapsedTime, conn);
            var ds4 = new DataSet();
            dataAdapter4.Fill(ds4);
            stagedDataGridView.DataSource = ds4.Tables[0];
            var elapsedTimeActual = stagedDataGridView.Rows[0].Cells[0].Value;

            //elapsed time average
            var elapsedTimeAverage = "";
            if (pendingRadioButton.Checked == true)
            {
                elapsedTimeAverage = " USE " + databaseSelect4.Text + " select ((sum(elapsedtime)/COUNT(*))/1000) / 60 as name from RunList rl left join rundet rd on rd.runlistno = rl.runlistno where rl.rectype='pay' and rd.ItemName = 'PayoutTypeNo' and rd.ItemValue = (select payouttypeno from PayoutType where payouttypeid = '" + payoutTypeSelect.Text + "')  and rl.finalizestatus='p'";
            }
            else if (finalizedRadioButton.Checked == true)
            {
                elapsedTimeAverage = " USE " + databaseSelect4.Text + " select ((sum(elapsedtime)/COUNT(*))/1000) / 60 as name from RunList rl left join rundet rd on rd.runlistno = rl.runlistno where rl.rectype='pay' and rd.ItemName = 'PayoutTypeNo' and rd.ItemValue = (select payouttypeno from PayoutType where payouttypeid = '" + payoutTypeSelect.Text + "')  and rl.finalizestatus='f'";
            }
            else if (reversedRadioButton.Checked == true)
            {
                elapsedTimeAverage = " USE " + databaseSelect4.Text + " select ((sum(elapsedtime)/COUNT(*))/1000) / 60 as name from RunList rl left join rundet rd on rd.runlistno = rl.runlistno where rl.rectype='pay' and rd.ItemName = 'PayoutTypeNo' and rd.ItemValue = (select payouttypeno from PayoutType where payouttypeid = '" + payoutTypeSelect.Text + "')  and rl.finalizestatus='r'";
            }
            var dataAdapter5 = new SqlDataAdapter(elapsedTimeAverage, conn);
            var ds5 = new DataSet();
            dataAdapter5.Fill(ds5);
            stagedDataGridView.DataSource = ds5.Tables[0];
            var elapsedTimeAverageActual = stagedDataGridView.Rows[0].Cells[0].Value;

            //fasterslower
            var fasterSlower = "";
            if (Convert.ToInt32(elapsedTimeActual) < Convert.ToInt32(elapsedTimeAverageActual))
            {
                fasterSlower = "faster";
            }
            else
            {
                fasterSlower = "slower";
            }

            //fasterslowerpercent
            decimal fasterSlowerPercent = 0;
            if (Convert.ToInt32(elapsedTimeActual) < Convert.ToInt32(elapsedTimeAverageActual))
            {
                fasterSlowerPercent = ((Convert.ToDecimal(elapsedTimeAverageActual) / Convert.ToDecimal(elapsedTimeActual)) - 1) * 100;
            }
            else
            {
                fasterSlowerPercent = fasterSlowerPercent = ((Convert.ToDecimal(elapsedTimeActual) / Convert.ToDecimal(elapsedTimeAverageActual)) - 1) * 100;
            }

            //task numbers
            var taskNumber = " USE " + databaseSelect4.Text + " select taskindex+1 as TaskNumber from runlist where RunListNoRoot=" + runListNo + " and TaskId is not null order by elapsedtime desc";
            var dataAdapter6 = new SqlDataAdapter(taskNumber, conn);
            var ds6 = new DataSet();
            dataAdapter6.Fill(ds6);
            stagedDataGridView.DataSource = ds6.Tables[0];
            var taskNumberArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

            //task ids
            var taskIds = " USE " + databaseSelect4.Text + " select taskid from runlist where RunListNoRoot=" + runListNo + " and TaskId is not null order by elapsedtime desc";
            var dataAdapter7 = new SqlDataAdapter(taskIds, conn);
            var ds7 = new DataSet();
            dataAdapter7.Fill(ds7);
            stagedDataGridView.DataSource = ds7.Tables[0];
            var taskIdsArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();



            var tasks = " USE " + databaseSelect4.Text + " select taskindex+1 as 'Task #',TaskId as 'Task Name',((sum(elapsedtime)/COUNT(*))/1000) / 60 as 'Task Run Time in Minutes' from runlist where RunListNoRoot=" + runListNo + " and TaskId is not null group by taskid, TaskIndex, ElapsedTime order by elapsedtime desc";
            var dataAdapter8 = new SqlDataAdapter(tasks, conn);
            var ds8 = new DataSet();
            dataAdapter8.Fill(ds8);
            benchmarkDataGridView.DataSource = ds8.Tables[0];

            benchmarkRichTextBox.Text = benchmarkRichTextBox.Text.Insert(0, Environment.NewLine +
                @"###########################################################################################" + System.Environment.NewLine +
                @"########################DataAnalysisTool - Payout Benchmark################################" + System.Environment.NewLine +
                @"###########################################################################################" + System.Environment.NewLine +
                @"Current Date: " + DateTime.Now + System.Environment.NewLine +
                @"Server: " + serverSelect4.Text + System.Environment.NewLine +
                @"Database: " + databaseSelect4.Text + System.Environment.NewLine +
                @"Payout Type: " + payoutTypeSelect.Text + System.Environment.NewLine +
                @"RunListNoRoot: " + runListNo +
                @"" + System.Environment.NewLine +
                @"" + System.Environment.NewLine +
                @"****************************************************" + System.Environment.NewLine +
                @"********************PAYOUT STATS********************" + System.Environment.NewLine +
                @"****************************************************" + System.Environment.NewLine +
                @"Elapsed time: " + elapsedTimeActual + " Minutes" + System.Environment.NewLine +
                @"Average payout time for the " + payoutTypeSelect.Text + " payout: " + elapsedTimeAverageActual + " Minutes" + System.Environment.NewLine +
                @"Percent " + fasterSlower + " than the payout average: " + fasterSlowerPercent + "%" + System.Environment.NewLine +
                @"" + System.Environment.NewLine +
                @"" + System.Environment.NewLine
                );
            progressBar1.MarqueeAnimationSpeed = 0;
            benchmarkProgressBar.Value = 100;
        }

        private void button28_Click(object sender, EventArgs e)
        {
            Process.Start(Application.UserAppDataPath + @"\Payout_Benchmarks");
        }

        private void benchmarkClearResults_Click(object sender, EventArgs e)
        {
            benchmarkRichTextBox.Clear();
        }

        private void benchmarkExportResults_Click(object sender, EventArgs e)
        {
            if (benchmarkRichTextBox.Text == null || benchmarkRichTextBox.Text == "")
            {
                MessageBox.Show("There are no results to export!", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            System.IO.Directory.CreateDirectory(Application.UserAppDataPath + @"\Payout_Benchmarks");
            string path = Application.UserAppDataPath + @"\Payout_Benchmarks\DataAnalysisTool_PB_Data_" + DateTime.Now.ToString("MM_dd_yyyy_HHmmss") + ".txt";
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
            {
                using (TextWriter tw = new StreamWriter(fs))
                {
                    for (int i = 0; i < benchmarkRichTextBox.Lines.Length; i++)
                    {
                        tw.WriteLine(benchmarkRichTextBox.Lines[i]);
                    }
                    // setup for export
                    benchmarkDataGridView.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
                    benchmarkDataGridView.SelectAll();
                    // hiding row headers to avoid extra \t in exported text
                    var rowHeaders = benchmarkDataGridView.RowHeadersVisible;
                    benchmarkDataGridView.RowHeadersVisible = false;

                    // ! creating text from grid values
                    string content = benchmarkDataGridView.GetClipboardContent().GetText();

                    // restoring grid state
                    benchmarkDataGridView.ClearSelection();
                    benchmarkDataGridView.RowHeadersVisible = rowHeaders;
                    tw.WriteLine(content);
                    tw.WriteLine("EOF.");
                }
            }
            importFormatProgressBar.Value = 90;
            importFormatProgressBar.Value = 100;
            MessageBox.Show("Payout Benchmark file has been created. \nLocation: " + path, "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            progressBar1.MarqueeAnimationSpeed = 0;
            Process.Start(path);
        }

        //*********************************************************************************************
        //*********************************/PAYOUT BENCHMARK TAB****************************************
        //*********************************************************************************************

        //*********************************************************************************************
        //*********************************API READINESS TAB*******************************************
        //*********************************************************************************************

        private void apiReadinessCheckButton_Click(object sender, EventArgs e)
        {

            apiReadinessProgressBar.Value = 0;
            apiReadinessProgressBar.Value = 10;

            //global vars
            progressBar1.MarqueeAnimationSpeed = 10;
            if (databaseSelect5.Text == "")
            {
                DialogResult result = MessageBox.Show("No database selected. \nPlease make sure you are connected to ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                importFormatProgressBar.Value = 0;
                return;
            }

            if (databaseSelect5.Text != "")
            {

                DialogResult result2 = MessageBox.Show("The DAT will check against the " + databaseSelect5.Text + " database.\nContinue?", "Data Analysis Tool", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                if (result2 == DialogResult.No)
                {
                    progressBar1.MarqueeAnimationSpeed = 0;
                    importFormatProgressBar.Value = 0;
                    return;
                }
            }

            apiUsersComboBox.Items.Clear();
            apiCallButton.Visible = false;
            apiUsersComboBox.Visible = false;
            apiUsersPictureBox.Visible = false;
            apiUsersPasswordPictureBox.Visible = false;
            apiUsersPasswordTextBox.Visible = false;
            apiRichTextBox.Clear();

            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect5.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();

            var secGroups = " USE " + databaseSelect5.Text + " select SecGroupId from secgroup where portalid=6 and prosta=1";
            var dataAdapter = new SqlDataAdapter(secGroups, conn);
            var ds = new DataSet();
            dataAdapter.Fill(ds);
            stagedDataGridView.DataSource = ds.Tables[0];
            var secGroupsArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                    .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

            var apiUsers = " USE " + databaseSelect5.Text + " select us.userid from UsrPortal up inner join UsrSet us on up.userno=us.userno where up.ProSta=1 and up.SecGroupNo in (select SecGroupNo from secgroup where portalid=6 and prosta=1)";
            var dataAdapter2 = new SqlDataAdapter(apiUsers, conn);
            var ds2 = new DataSet();
            dataAdapter2.Fill(ds2);
            stagedDataGridView.DataSource = ds2.Tables[0];
            var apiUsersArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                    .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

            var apiEnabled = " USE " + databaseSelect5.Text + " select case when enabled=1 then 'Yes' else 'No' end as 'Enabled' from feature where FeatureId='System API''s'";
            var dataAdapter3 = new SqlDataAdapter(apiEnabled, conn);
            var ds3 = new DataSet();
            dataAdapter3.Fill(ds3);
            stagedDataGridView.DataSource = ds3.Tables[0];
            var apiEnabledFinal = stagedDataGridView.Rows[0].Cells[0].Value;

            conn.Close();

            progressBar1.MarqueeAnimationSpeed = 0;

            apiRichTextBox.AppendText(Environment.NewLine +
                @"###########################################################################################" + System.Environment.NewLine +
                @"########################DataAnalysisTool - API Readiness###################################" + System.Environment.NewLine +
                @"###########################################################################################" + System.Environment.NewLine +
                @"Current Date: " + DateTime.Now + System.Environment.NewLine +
                @"Server: " + serverSelect5.Text + System.Environment.NewLine +
                @"Database: " + databaseSelect5.Text + System.Environment.NewLine +
                @"" + System.Environment.NewLine +
                @"" + System.Environment.NewLine +
                @"****************************************************" + System.Environment.NewLine +
                @"********************RUN RESULTS*********************" + System.Environment.NewLine +
                @"****************************************************" + System.Environment.NewLine
                );

            apiRichTextBox.AppendText(@"" + System.Environment.NewLine + "API enabled: " + System.Environment.NewLine + apiEnabledFinal);

            apiRichTextBox.AppendText(@"" + System.Environment.NewLine);

            if (apiEnabledFinal.Equals("Yes"))
            {
                apiPictureBox.Image = Properties.Resources.greenCheck;
            }
            else
            {
                apiRichTextBox.AppendText(Environment.NewLine + @"Please enable API's within the Global Features.");
                apiPictureBox.Image = Properties.Resources.global;
                return;
            }

            if (secGroupsArray.Length == 0)
            {
                apiRichTextBox.AppendText(Environment.NewLine + @"Please enable or create an API security group.");
                apiPictureBox.Image = Properties.Resources.sec;
                return;
            }
            else
            {
                apiPictureBox.Image = Properties.Resources.greenCheck;
            }

            apiRichTextBox.AppendText(Environment.NewLine + @"API Security Groups:");
            foreach (var sec in secGroupsArray)
            {
                apiRichTextBox.AppendText(@"" + System.Environment.NewLine + sec);
            }

            apiRichTextBox.AppendText(Environment.NewLine + @"");

            apiRichTextBox.AppendText(Environment.NewLine + @"Optionally, configure the system.api.ip.whitelist to restrict access to a range of client IP addresses. 
(Admin > Configuration > Options) E.g. restrict access to internal IP addresses. Note that if this option
is not configured, the System API's may be accessed from any IP address. This may be considered a security 
risk if your ICM instance is externally accessible.");

            apiRichTextBox.AppendText(Environment.NewLine + @"");
            apiRichTextBox.AppendText(Environment.NewLine + @"API Users:");

            if (apiUsersArray.Length >= 1)
            {
                foreach (var api in apiUsersArray)
                {
                    apiRichTextBox.AppendText(@"" + System.Environment.NewLine + api);
                }
                apiCallButton.Visible = true;
                apiUsersComboBox.Visible = true;
                apiUsersPictureBox.Visible = true;
                apiUsersPasswordPictureBox.Visible = true;
                apiUsersPasswordTextBox.Visible = true;
                apiEnvLabel1.Visible = true;
                apiEnvLabel2.Visible = true;
                apiEnvLabelMain.Visible = true;
                for (int i = 0; i < apiUsersArray.Length; i++)
                {
                    apiUsersComboBox.Items.Add(apiUsersArray[i]);
                }
                apiRichTextBox.AppendText(Environment.NewLine + @"");
                apiRichTextBox.AppendText(Environment.NewLine + @"It looks like this environment is ready to test an API call. If you would like to do this, please select a user above, type the password, then click Test Call");
                apiCallButton.Visible = true;
            }
            else
            {
                apiRichTextBox.AppendText(Environment.NewLine + @"No API users found. Define one or more Users with access to the '''System API's''' AppArea. (Admin > Security > Users).");
                apiPictureBox.Image = Properties.Resources.apiuser;
                return;
            }
            apiRichTextBox.AppendText(Environment.NewLine + @"");
            apiReadinessProgressBar.Value = 100;
        }

        private void apiCallButton_Click(object sender, EventArgs e)
        {
            apiRichTextBox.Clear();
            using (var client = new HttpClient(new HttpClientHandler { AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate }))
            {
                apiRichTextBox.AppendText(Environment.NewLine +
                    @"###########################################################################################" + System.Environment.NewLine +
                    @"########################DataAnalysisTool - API Readiness###################################" + System.Environment.NewLine +
                    @"###########################################################################################" + System.Environment.NewLine +
                    @"Current Date: " + DateTime.Now + System.Environment.NewLine +
                    @"Server: " + serverSelect5.Text + System.Environment.NewLine +
                    @"Database: " + databaseSelect5.Text + System.Environment.NewLine +
                    @"" + System.Environment.NewLine +
                    @"" + System.Environment.NewLine +
                    @"****************************************************" + System.Environment.NewLine +
                    @"********************API CALL RESULTS****************" + System.Environment.NewLine +
                    @"****************************************************" + System.Environment.NewLine
                    );
                client.BaseAddress = new Uri("https://welltest2.callidusinsurance.net/ICM/REST/auth/login?u=" + apiUsersComboBox.Text + "&p=" + apiUsersPasswordTextBox.Text);
                HttpResponseMessage response = client.GetAsync("").Result;
                response.EnsureSuccessStatusCode();
                string result = response.Content.ReadAsStringAsync().Result;
                Console.WriteLine("Result: " + result);
                apiRichTextBox.AppendText(@"" + System.Environment.NewLine + result);
            }
        }

        private void aPILogFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(Application.UserAppDataPath + @"\API_Readiness_Check");
        }

        //*********************************************************************************************
        //*********************************/API READINESS TAB*******************************************
        //*********************************************************************************************
        private void toolStripStatusLabel4_Click(object sender, EventArgs e)
        {
            ifRowCounterToolStripStatusLabel.Text = importedfileDataGridView.Rows.Count.ToString();
        }

        public DataTable ReadTxtComma(string fileName)
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

        private void pipeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "TXT | *.txt", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        importFormatProgressBar.Value = 20;
                        importedfileDataGridView.DataSource = ReadTxtPipe(ofd.FileName);
                        importFormatActualFileNameToolStripStatusLabel.Text = ofd.FileName;
                        importFormatActualFileNameToolStripStatusLabel.Visible = true;
                        ifRowCounterToolStripStatusLabel.Text = importedfileDataGridView.Rows.Count.ToString();
                        ifRowCountLabelToolStripStatusLabel.Visible = true;
                        ifRowCounterToolStripStatusLabel.Visible = true;
                        seperator3ToolStripStatusLabel.Visible = true;
                        importFormatFileNameToolStripStatusLabel.Visible = true;
                        systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading TXT: " + ofd.FileName + "...Done.");
                    }
                    else
                    {
                        importFormatProgressBar.Value = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                importFormatProgressBar.Value = 0;
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            importFormatProgressBar.Value = 100;
            progressBar1.MarqueeAnimationSpeed = 0;
        }

        public DataTable ReadTxtPipe(string fileName)
        {
            importFormatProgressBar.Value = 30;
            DataTable dt = new DataTable();
            string[] columns = null;

            var lines = File.ReadAllLines(fileName);

            if (importformatIncludeHeaderRowButton.Checked == false)
            {
                importFormatProgressBar.Value = 50;
                if (lines.Count() > 0)
                {
                    importFormatProgressBar.Value = 60;
                    columns = lines[0].Split(new char[] { '|' });
                }

                int columnCount1 = columns.Count();
                for (int i = 0; i < columnCount1; i++)
                {
                    dt.Columns.Add("column " + (i+1));
                }

                // reading rest of the data
                for (int i = 0; i < lines.Count(); i++)
                {
                    DataRow dr = dt.NewRow();
                    string[] values = lines[i].Split(new char[] { '|' });

                    for (int j = 0; j < values.Count() && j < columns.Count(); j++)
                        dr[j] = values[j];

                    dt.Rows.Add(dr);
                }
                importFormatProgressBar.Value = 70;
                return dt;
            }
            else
            {
                importFormatProgressBar.Value = 50;
                if (lines.Count() > 0)
                {
                    importFormatProgressBar.Value = 60;
                    columns = lines[0].Split(new char[] { '|' });

                    foreach (var column in columns)
                        dt.Columns.Add(column);
                }

                // reading rest of the data
                for (int i = 1; i < lines.Count(); i++)
                {
                    DataRow dr = dt.NewRow();
                    string[] values = lines[i].Split(new char[] { '|' });
                    for (int j = 0; j < values.Count() && j < columns.Count(); j++)
                        dr[j] = values[j];
                    dt.Rows.Add(dr);
                }
                importFormatProgressBar.Value = 70;
                return dt;
            }
        }

        private void dateComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //day check
            if (dateComboBox1.Text == "d" || dateComboBox1.Text == "dd" || dateComboBox1.Text == "ddd" || dateComboBox1.Text == "dddd")
            {
                if (dateComboBox2.Text == "d" || dateComboBox3.Text == "d")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "dd" || dateComboBox3.Text == "dd")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "ddd" || dateComboBox3.Text == "ddd")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "dddd" || dateComboBox3.Text == "dddd")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
            }

            //month check
            if (dateComboBox1.Text == "m" || dateComboBox1.Text == "mm" || dateComboBox1.Text == "M" || dateComboBox1.Text == "MM" || dateComboBox1.Text == "MMM" || dateComboBox1.Text == "MMM" || dateComboBox1.Text == "MMMM")
            {
                if (dateComboBox2.Text == "m" || dateComboBox3.Text == "m")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "mm" || dateComboBox3.Text == "mm")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "M" || dateComboBox3.Text == "M")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "MM" || dateComboBox3.Text == "MM")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "MMM" || dateComboBox3.Text == "MMM")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "MMMM" || dateComboBox3.Text == "MMMM")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
            }

            //year check
            if (dateComboBox1.Text == "y" || dateComboBox1.Text == "yy" || dateComboBox1.Text == "yyyy")
            {
                if (dateComboBox2.Text == "y" || dateComboBox3.Text == "y")
                {
                    MessageBox.Show("Cannot have more than one 'year' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "yy" || dateComboBox3.Text == "yy")
                {
                    MessageBox.Show("Cannot have more than one 'year' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "yyyy" || dateComboBox3.Text == "yyyy")
                {
                    MessageBox.Show("Cannot have more than one 'year' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
            }


            dateFormat.Text = "Date Format: "+dateComboBox1.Text+ dateComboBoxSeperator.Text + dateComboBox2.Text+ dateComboBoxSeperator.Text+dateComboBox3.Text;
        }

        private void dateComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //day check
            if (dateComboBox2.Text == "d" || dateComboBox2.Text == "dd" || dateComboBox2.Text == "ddd" || dateComboBox2.Text == "dddd")
            {
                if (dateComboBox1.Text == "d" || dateComboBox3.Text == "d")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "dd" || dateComboBox3.Text == "dd")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "ddd" || dateComboBox3.Text == "ddd")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "dddd" || dateComboBox3.Text == "dddd")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
            }

            //month check
            if (dateComboBox2.Text == "m" || dateComboBox2.Text == "mm" || dateComboBox2.Text == "M" || dateComboBox2.Text == "MM" || dateComboBox2.Text == "MMM" || dateComboBox2.Text == "MMM" || dateComboBox2.Text == "MMMM")
            {
                if (dateComboBox1.Text == "m" || dateComboBox3.Text == "m")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "mm" || dateComboBox3.Text == "mm")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "M" || dateComboBox3.Text == "M")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "MM" || dateComboBox3.Text == "MM")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "MMM" || dateComboBox3.Text == "MMM")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "MMMM" || dateComboBox3.Text == "MMMM")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
            }

            //year check
            if (dateComboBox2.Text == "y" || dateComboBox2.Text == "yy" || dateComboBox2.Text == "yyyy")
            {
                if (dateComboBox1.Text == "y" || dateComboBox3.Text == "y")
                {
                    MessageBox.Show("Cannot have more than one 'year' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "yy" || dateComboBox3.Text == "yy")
                {
                    MessageBox.Show("Cannot have more than one 'year' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "yyyy" || dateComboBox3.Text == "yyyy")
                {
                    MessageBox.Show("Cannot have more than one 'year' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
            }
            dateFormat.Text = "Date Format: " + dateComboBox1.Text + dateComboBoxSeperator.Text + dateComboBox2.Text + dateComboBoxSeperator.Text + dateComboBox3.Text;
        }

        private void dateComboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            //day check
            if (dateComboBox3.Text == "d" || dateComboBox3.Text == "dd" || dateComboBox3.Text == "ddd" || dateComboBox3.Text == "dddd")
            {
                if (dateComboBox1.Text == "d" || dateComboBox2.Text == "d")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "dd" || dateComboBox2.Text == "dd")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "ddd" || dateComboBox2.Text == "ddd")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "dddd" || dateComboBox2.Text == "dddd")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
            }

            //month check
            if (dateComboBox3.Text == "m" || dateComboBox3.Text == "mm" || dateComboBox3.Text == "M" || dateComboBox3.Text == "MM" || dateComboBox3.Text == "MMM" || dateComboBox3.Text == "MMM" || dateComboBox3.Text == "MMMM")
            {
                if (dateComboBox1.Text == "m" || dateComboBox2.Text == "m")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "mm" || dateComboBox2.Text == "mm")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "M" || dateComboBox2.Text == "M")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "MM" || dateComboBox2.Text == "MM")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "MMM" || dateComboBox2.Text == "MMM")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "MMMM" || dateComboBox2.Text == "MMMM")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
            }

            //year check
            if (dateComboBox3.Text == "y" || dateComboBox3.Text == "yy" || dateComboBox3.Text == "yyyy")
            {
                if (dateComboBox2.Text == "y" || dateComboBox1.Text == "y")
                {
                    MessageBox.Show("Cannot have more than one 'year' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "yy" || dateComboBox1.Text == "yy")
                {
                    MessageBox.Show("Cannot have more than one 'year' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "yyyy" || dateComboBox1.Text == "yyyy")
                {
                    MessageBox.Show("Cannot have more than one 'year' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
            }
            dateFormat.Text = "Date Format: " + dateComboBox1.Text + dateComboBoxSeperator.Text + dateComboBox2.Text + dateComboBoxSeperator.Text + dateComboBox3.Text;
        }

        private void dateComboBoxSeperator_SelectedIndexChanged(object sender, EventArgs e)
        {
            dateFormat.Text = "Date Format: " + dateComboBox1.Text + dateComboBoxSeperator.Text + dateComboBox2.Text + dateComboBoxSeperator.Text + dateComboBox3.Text;
        }

        private void button25_Click(object sender, EventArgs e)
        {
            try
            {
                int length = int.Parse(importFormatJumpToRowTextBox.Text);
                importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[length - 1].Cells[0];
                importedfileDataGridView.Rows[length - 1].Selected = true;
            }
            catch { MessageBox.Show("That column does not exist!", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1); }
        }

        private void checkBox4_Click(object sender, EventArgs e)
        {
            if (databaseSelect.Text != "")
            {

                int value = databaseSelect.SelectedIndex;
                databaseSelect.SelectedIndex = -1;
                databaseSelect.SelectedIndex = value;
            }
        }

        //------------------EXIT APP ACTION END------------------------------------------------------
        /*
         * ############################################################################################   
         * ############################################################################################
         * ####################PRODUCTION CODE END#####################################################
         * ############################################################################################
         * ############################################################################################
        */

        private void button27_Click(object sender, EventArgs e)
        {
            SqlConnection pubsConn = new SqlConnection(@"Data Source = " + serverSelect5.Text + "; Initial Catalog = master; Integrated Security = True");
            SqlCommand logoCMD = new SqlCommand(" USE " + databaseSelect5.Text + " select content from outfile where runlistno =15408457951330000", pubsConn);

            FileStream fs;                          // Writes the BLOB to a file (*.bmp).
            BinaryWriter bw;                        // Streams the BLOB to the FileStream object.

            int bufferSize = 100;                   // Size of the BLOB buffer.
            byte[] outbyte = new byte[bufferSize];  // The BLOB byte[] buffer to be filled by GetBytes.
            long retval;                            // The bytes returned from GetBytes.
            long startIndex = 0;                    // The starting position in the BLOB output.

            string pub_id = "";                     // The publisher id to use in the file name.

            // Open the connection and read data into the DataReader.
            pubsConn.Open();
            SqlDataReader myReader = logoCMD.ExecuteReader(CommandBehavior.SequentialAccess);

            while (myReader.Read())
            {
                // Get the publisher id, which must occur before getting the logo.
                // Create a file to hold the output.
                fs = new FileStream("icmlog" + pub_id + ".log", FileMode.OpenOrCreate, FileAccess.Write);
                bw = new BinaryWriter(fs);
                // Reset the starting byte for the new BLOB.
                MessageBox.Show("" + startIndex);
                startIndex = 0;
                // Read the bytes into outbyte[] and retain the number of bytes returned.
                MessageBox.Show("outbyte" + outbyte);
                MessageBox.Show("buffersize" + bufferSize);
                retval = myReader.GetBytes(1, startIndex, outbyte, 0, bufferSize);//fails
                MessageBox.Show("1859");
                // Continue reading and writing while there are bytes beyond the size of the buffer.
                while (retval == bufferSize)
                {
                    bw.Write(outbyte);
                    bw.Flush();

                    // Reposition the start index to the end of the last buffer and fill the buffer.
                    startIndex += bufferSize;
                    retval = myReader.GetBytes(1, startIndex, outbyte, 0, bufferSize);
                }

                // Write the remaining buffer.
                bw.Write(outbyte, 0, (int)retval - 1);
                bw.Flush();

                // Close the output file.
                bw.Close();
                fs.Close();
            }

            // Close the reader and the connection.
            myReader.Close();
            pubsConn.Close();
        }



        private void copyAlltoClipboard()
        {
            importedfileDataGridView.SelectAll();
            DataObject dataObj = importedfileDataGridView.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }



        private void fromDateEnableCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (fromDateEnableCheckBox.Checked == true)
            {
                dateYearFromTextBox.ReadOnly = false;
                dateMonthFromTextBox.ReadOnly = false;
                dateDayFromTextBox.ReadOnly = false;
            }

            else
            {
                dateYearFromTextBox.ReadOnly = true;
                dateMonthFromTextBox.ReadOnly = true;
                dateDayFromTextBox.ReadOnly = true;
            }
        }

        private void toDateEnableCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (toDateEnableCheckBox.Checked == true)
            {
                dateYearToTextBox.ReadOnly = false;
                dateMonthToTextBox.ReadOnly = false;
                dateDayToTextBox.ReadOnly = false;
            }

            else
            {
                dateYearToTextBox.ReadOnly = true;
                dateMonthToTextBox.ReadOnly = true;
                dateDayToTextBox.ReadOnly = true;
            }
        }

        private void dateTextBox_KeyUp(object sender, KeyEventArgs e)
        {
            dateRangeLabel.Text = "Date Range: " + dateYearFromTextBox.Text + dateMonthFromTextBox.Text + dateDayFromTextBox.Text + " - " + dateYearToTextBox.Text + dateMonthToTextBox.Text + dateDayToTextBox.Text;
        }

        private void envChangesGoPictureBox_Click(object sender, EventArgs e)
        {
            envChangesProgressBar.Value = 0;
            envChangesProgressBar.Value = 10;

            //global vars
            progressBar1.MarqueeAnimationSpeed = 10;


            if (databaseSelect6.Text == "")
            {
                DialogResult result = MessageBox.Show("No database selected. \nPlease make sure you are connected to ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                envChangesProgressBar.Value = 0;
                return;
            }

            if (userIDTextBox.Text == "")
            {
                DialogResult result = MessageBox.Show("No UserID entered. \nPlease enter a UserID", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                envChangesProgressBar.Value = 0;
                return;
            }
            var fromDate = dateYearFromTextBox.Text + dateMonthFromTextBox.Text + dateDayFromTextBox.Text;
            var toDate = dateYearToTextBox.Text + dateMonthToTextBox.Text + dateDayToTextBox.Text;
            if (fromDateEnableCheckBox.Checked == true && fromDate.Length != 8)
            {
                MessageBox.Show("Incorrect date format on the From Section. \nPlease make sure you are using YYYYMMDD", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                envChangesProgressBar.Value = 0;
                return;
            }
            if (toDateEnableCheckBox.Checked == true && toDate.Length != 8)
            {
                MessageBox.Show("Incorrect date format on the From Section. \nPlease make sure you are using YYYYMMDD", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                envChangesProgressBar.Value = 0;
                return;
            }
            envChangesRichTextBox.Clear();
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect6.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            envChangesRichTextBox.AppendText(Environment.NewLine +
                @"###########################################################################################" + System.Environment.NewLine +
                @"########################DataAnalysisTool - Environment Changes#############################" + System.Environment.NewLine +
                @"###########################################################################################" + System.Environment.NewLine +
                @"Current Date: " + DateTime.Now + System.Environment.NewLine +
                @"Server: " + serverSelect6.Text + System.Environment.NewLine +
                @"Database: " + databaseSelect6.Text + System.Environment.NewLine +
                @"User: " + userIDTextBox.Text + System.Environment.NewLine +
                @"" + dateRangeLabel.Text + System.Environment.NewLine +
                @"" + System.Environment.NewLine +
                @"" + System.Environment.NewLine +
                @"****************************************************" + System.Environment.NewLine +
                @"********************RUN RESULTS*********************" + System.Environment.NewLine +
                @"****************************************************" + System.Environment.NewLine
                );

            //Import Formats
            envChangesRichTextBox.AppendText(@"" + System.Environment.NewLine);
            if (envChangesCheckBox1.Checked == true)
            {
                var changedImportformats = "";
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8 && toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedImportformats = " USE " + databaseSelect6.Text + " select importformatid from importformat where lstuser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8)
                {
                    changedImportformats = " USE " + databaseSelect6.Text + " select importformatid from importformat where lstuser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate;
                }
                if (toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedImportformats = " USE " + databaseSelect6.Text + " select importformatid from importformat where lstuser=" + "'" + userIDTextBox.Text + "'" + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == false && toDateEnableCheckBox.Checked == false)
                {
                    changedImportformats = " USE " + databaseSelect6.Text + " select importformatid from importformat where lstuser=" + "'" + userIDTextBox.Text + "'";
                }
                var dataAdapter = new SqlDataAdapter(changedImportformats, conn);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                stagedDataGridView.DataSource = ds.Tables[0];
                var changedImportFormatsArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                envChangesRichTextBox.AppendText(Environment.NewLine + @"Changed Import Formats:");
                foreach (var sec in changedImportFormatsArray)
                {
                    envChangesRichTextBox.AppendText(@"" + System.Environment.NewLine + sec);
                }
            }
            //Expressions
            envChangesRichTextBox.AppendText(Environment.NewLine + @"");
            if (envChangesCheckBox2.Checked == true)
            {
                var changedExpressions = "";
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8 && toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedExpressions = " USE " + databaseSelect6.Text + " select expressionid from expression where lstuser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8)
                {
                    changedExpressions = " USE " + databaseSelect6.Text + " select expressionid from expression where lstuser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate;
                }
                if (toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedExpressions = " USE " + databaseSelect6.Text + " select expressionid from expression where lstuser=" + "'" + userIDTextBox.Text + "'" + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == false && toDateEnableCheckBox.Checked == false)
                {
                    changedExpressions = " USE " + databaseSelect6.Text + " select expressionid from expression where lstuser=" + "'" + userIDTextBox.Text + "'";
                }
                var dataAdapter = new SqlDataAdapter(changedExpressions, conn);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                stagedDataGridView.DataSource = ds.Tables[0];
                var changedExpressionsArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                envChangesRichTextBox.AppendText(Environment.NewLine + @"Changed Expressions:");
                foreach (var sec in changedExpressionsArray)
                {
                    envChangesRichTextBox.AppendText(@"" + System.Environment.NewLine + sec);
                }
            }
            //QBQ
            envChangesRichTextBox.AppendText(Environment.NewLine + @"");
            if (envChangesCheckBox3.Checked == true)
            {
                var changedQBQ = "";
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8 && toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedQBQ = " USE " + databaseSelect6.Text + " select QBQueryId from QBQuery where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >"+fromDate+" and lstdate < "+toDate;
                }
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8)
                {
                    changedQBQ = " USE " + databaseSelect6.Text + " select QBQueryId from QBQuery where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate;
                }
                if (toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedQBQ = " USE " + databaseSelect6.Text + " select QBQueryId from QBQuery where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == false && toDateEnableCheckBox.Checked == false)
                {
                    changedQBQ = " USE " + databaseSelect6.Text + " select QBQueryId from QBQuery where LstUser=" + "'" + userIDTextBox.Text + "'";
                }
                
                var dataAdapter = new SqlDataAdapter(changedQBQ, conn);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                stagedDataGridView.DataSource = ds.Tables[0];
                var changedQBQArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                envChangesRichTextBox.AppendText(Environment.NewLine + @"Changed QBQ:");
                foreach (var sec in changedQBQArray)
                {
                    envChangesRichTextBox.AppendText(@"" + System.Environment.NewLine + sec);
                }
            }
            //Xref
            envChangesRichTextBox.AppendText(Environment.NewLine + @"");
            if (envChangesCheckBox4.Checked == true)
            {
                var changedXref = "";
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8 && toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedXref = " USE " + databaseSelect6.Text + " select ExtCrossRefTypeId from ExtCrossRefType where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8)
                {
                    changedXref = " USE " + databaseSelect6.Text + " select ExtCrossRefTypeId from ExtCrossRefType where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate;
                }
                if (toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedXref = " USE " + databaseSelect6.Text + " select ExtCrossRefTypeId from ExtCrossRefType where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == false && toDateEnableCheckBox.Checked == false)
                {
                    changedXref = " USE " + databaseSelect6.Text + " select ExtCrossRefTypeId from ExtCrossRefType where LstUser=" + "'" + userIDTextBox.Text + "'";
                }
                var dataAdapter = new SqlDataAdapter(changedXref, conn);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                stagedDataGridView.DataSource = ds.Tables[0];
                var changedImportFormatsArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                envChangesRichTextBox.AppendText(Environment.NewLine + @"Changed Cross-Refs:");
                foreach (var sec in changedImportFormatsArray)
                {
                    envChangesRichTextBox.AppendText(@"" + System.Environment.NewLine + sec);
                }
            }
            //Field Default
            envChangesRichTextBox.AppendText(Environment.NewLine + @"");
            if (envChangesCheckBox5.Checked == true)
            {
                var changedFieldDefault = "";
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8 && toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedFieldDefault = " USE " + databaseSelect6.Text + " select 'EntName: '+EntName+' FldName: '+FldName from FieldDefault where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8)
                {
                    changedFieldDefault = " USE " + databaseSelect6.Text + " select 'EntName: '+EntName+' FldName: '+FldName from FieldDefault where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate;
                }
                if (toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedFieldDefault = " USE " + databaseSelect6.Text + " select 'EntName: '+EntName+' FldName: '+FldName from FieldDefault where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == false && toDateEnableCheckBox.Checked == false)
                {
                    changedFieldDefault = " USE " + databaseSelect6.Text + " select 'EntName: '+EntName+' FldName: '+FldName from FieldDefault where LstUser=" + "'" + userIDTextBox.Text + "'";
                }
                var dataAdapter = new SqlDataAdapter(changedFieldDefault, conn);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                stagedDataGridView.DataSource = ds.Tables[0];
                var changedImportFormatsArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                envChangesRichTextBox.AppendText(Environment.NewLine + @"Changed Field Defaults:");
                foreach (var sec in changedImportFormatsArray)
                {
                    envChangesRichTextBox.AppendText(@"" + System.Environment.NewLine + sec);
                }
            }
            //BEU
            envChangesRichTextBox.AppendText(Environment.NewLine + @"");
            if (envChangesCheckBox6.Checked == true)
            {
                var changedBEU = "";
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8 && toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedBEU = " USE " + databaseSelect6.Text + " select BatchEntityUpdateId from BatchEntityUpdate where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8)
                {
                    changedBEU = " USE " + databaseSelect6.Text + " select BatchEntityUpdateId from BatchEntityUpdate where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate;
                }
                if (toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedBEU = " USE " + databaseSelect6.Text + " select BatchEntityUpdateId from BatchEntityUpdate where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == false && toDateEnableCheckBox.Checked == false)
                {
                    changedBEU = " USE " + databaseSelect6.Text + " select BatchEntityUpdateId from BatchEntityUpdate where LstUser=" + "'" + userIDTextBox.Text + "'";
                }
                var dataAdapter = new SqlDataAdapter(changedBEU, conn);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                stagedDataGridView.DataSource = ds.Tables[0];
                var changedImportFormatsArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                envChangesRichTextBox.AppendText(Environment.NewLine + @"Changed BEUs:");
                foreach (var sec in changedImportFormatsArray)
                {
                    envChangesRichTextBox.AppendText(@"" + System.Environment.NewLine + sec);
                }
            }
            //Report Forms
            envChangesRichTextBox.AppendText(Environment.NewLine + @"");
            if (envChangesCheckBox6.Checked == true)
            {
                var changedReportForms = "";
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8 && toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedReportForms = " USE " + databaseSelect6.Text + " select Name from ReportForm where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8)
                {
                    changedReportForms = " USE " + databaseSelect6.Text + " select Name from ReportForm where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate;
                }
                if (toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedReportForms = " USE " + databaseSelect6.Text + " select Name from ReportForm where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == false && toDateEnableCheckBox.Checked == false)
                {
                    changedReportForms = " USE " + databaseSelect6.Text + " select Name from ReportForm where LstUser=" + "'" + userIDTextBox.Text + "'";
                }
                var dataAdapter = new SqlDataAdapter(changedReportForms, conn);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                stagedDataGridView.DataSource = ds.Tables[0];
                var changedImportFormatsArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                envChangesRichTextBox.AppendText(Environment.NewLine + @"Changed Report Forms:");
                foreach (var sec in changedImportFormatsArray)
                {
                    envChangesRichTextBox.AppendText(@"" + System.Environment.NewLine + sec);
                }
            }
            //Report Templates
            envChangesRichTextBox.AppendText(Environment.NewLine + @"");
            if (envChangesCheckBox6.Checked == true)
            {
                var changedReportTemplates = "";
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8 && toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedReportTemplates = " USE " + databaseSelect6.Text + " select ReportId from JasperReport where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8)
                {
                    changedReportTemplates = " USE " + databaseSelect6.Text + " select ReportId from JasperReport where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate;
                }
                if (toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedReportTemplates = " USE " + databaseSelect6.Text + " select ReportId from JasperReport where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == false && toDateEnableCheckBox.Checked == false)
                {
                    changedReportTemplates = " USE " + databaseSelect6.Text + " select ReportId from JasperReport where LstUser=" + "'" + userIDTextBox.Text + "'";
                }
                var dataAdapter = new SqlDataAdapter(changedReportTemplates, conn);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                stagedDataGridView.DataSource = ds.Tables[0];
                var changedImportFormatsArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                envChangesRichTextBox.AppendText(Environment.NewLine + @"Changed Report Templates:");
                foreach (var sec in changedImportFormatsArray)
                {
                    envChangesRichTextBox.AppendText(@"" + System.Environment.NewLine + sec);
                }
            }
            envChangesRichTextBox.AppendText(Environment.NewLine + @"");
            envChangesProgressBar.Value = 100;
        }



        private void apiExportResultsPictureBox_Click(object sender, EventArgs e)
        {
            if (apiRichTextBox.Text == null || apiRichTextBox.Text == "")
            {
                MessageBox.Show("There are no results to export!", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            System.IO.Directory.CreateDirectory(Application.UserAppDataPath + @"\API_Readiness_Check");
            string path = Application.UserAppDataPath + @"\API_Readiness_Check\DataAnalysisTool_API_Check_" + DateTime.Now.ToString("MM_dd_yyyy_HHmmss") + ".txt";
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
            {
                using (TextWriter tw = new StreamWriter(fs))
                {
                    for (int i = 0; i < apiRichTextBox.Lines.Length; i++)
                    {
                        tw.WriteLine(apiRichTextBox.Lines[i]);
                    }
                    tw.WriteLine("EOF.");
                }
            }
            apiReadinessProgressBar.Value = 90;
            apiReadinessProgressBar.Value = 100;
            MessageBox.Show("API Readiness file has been created. \nLocation: " + path, "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            progressBar1.MarqueeAnimationSpeed = 0;
            Process.Start(path);
        }

        private void apiClearResultsPictureBox_Click(object sender, EventArgs e)
        {
            apiRichTextBox.Clear();
        }

        private void benchmarkExportResultsPictureBox_Click(object sender, EventArgs e)
        {
            if (benchmarkRichTextBox.Text == null || benchmarkRichTextBox.Text == "")
            {
                MessageBox.Show("There are no results to export!", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            System.IO.Directory.CreateDirectory(Application.UserAppDataPath + @"\Payout_Benchmarks");
            string path = Application.UserAppDataPath + @"\Payout_Benchmarks\DataAnalysisTool_PB_Data_" + DateTime.Now.ToString("MM_dd_yyyy_HHmmss") + ".txt";
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
            {
                using (TextWriter tw = new StreamWriter(fs))
                {
                    for (int i = 0; i < benchmarkRichTextBox.Lines.Length; i++)
                    {
                        tw.WriteLine(benchmarkRichTextBox.Lines[i]);
                    }
                    // setup for export
                    benchmarkDataGridView.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
                    benchmarkDataGridView.SelectAll();
                    // hiding row headers to avoid extra \t in exported text
                    var rowHeaders = benchmarkDataGridView.RowHeadersVisible;
                    benchmarkDataGridView.RowHeadersVisible = false;

                    // ! creating text from grid values
                    string content = benchmarkDataGridView.GetClipboardContent().GetText();

                    // restoring grid state
                    benchmarkDataGridView.ClearSelection();
                    benchmarkDataGridView.RowHeadersVisible = rowHeaders;
                    tw.WriteLine(content);
                    tw.WriteLine("EOF.");
                }
            }
            importFormatProgressBar.Value = 90;
            importFormatProgressBar.Value = 100;
            MessageBox.Show("Payout Benchmark file has been created. \nLocation: " + path, "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            progressBar1.MarqueeAnimationSpeed = 0;
            Process.Start(path);
        }

        private void benchmarkClearResultsPictureBox_Click(object sender, EventArgs e)
        {
            benchmarkRichTextBox.Clear();
        }

        private void sqlQueryGoPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.sqlQueryGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go3));
        }

        private void sqlQueryGoPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.sqlQueryGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.sqlQueryGoPictureBox, "Run the tool!");
        }

        private void sqlQueryGoPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.sqlQueryGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.sqlQueryGoPictureBox, "Run the tool!");
        }

        private void sqlQueryGoPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.sqlQueryGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
        }

        private void exportResultsPictureBox_Click(object sender, EventArgs e)
        {
            if (envChangesRichTextBox.Text == null || apiRichTextBox.Text == "")
            {
                MessageBox.Show("There are no results to export!", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            System.IO.Directory.CreateDirectory(Application.UserAppDataPath + @"\API_Readiness_Check");
            string path = Application.UserAppDataPath + @"\API_Readiness_Check\DataAnalysisTool_API_Check_" + DateTime.Now.ToString("MM_dd_yyyy_HHmmss") + ".txt";
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
            {
                using (TextWriter tw = new StreamWriter(fs))
                {
                    for (int i = 0; i < apiRichTextBox.Lines.Length; i++)
                    {
                        tw.WriteLine(apiRichTextBox.Lines[i]);
                    }
                    tw.WriteLine("EOF.");
                }
            }
            apiReadinessProgressBar.Value = 90;
            apiReadinessProgressBar.Value = 100;
            MessageBox.Show("Environment changes file has been created. \nLocation: " + path, "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            progressBar1.MarqueeAnimationSpeed = 0;
            Process.Start(path);
        }

        private void openInExcelPictureBox_Click(object sender, EventArgs e)
        {
            copyAlltoClipboard();
            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Microsoft.Office.Interop.Excel.Application();
            xlexcel.Visible = true;
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
            CR.Select();
            xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
        }

        private void openInExcelPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.openInExcelPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_open_in_excel3));
        }

        private void openInExcelPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.openInExcelPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_open_in_excel2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.openInExcelPictureBox, "Open the table in Excel.");
        }

        private void openInExcelPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.openInExcelPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_open_in_excel));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.openInExcelPictureBox, "Open the table in Excel.");
        }

        private void openInExcelPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.openInExcelPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_open_in_excel));
        }

        private void legendButtonPictureBox_Click(object sender, EventArgs e)
        {
            DataGridViewLegend legend = new DataGridViewLegend();

            while (Application.OpenForms.Count > 1)
            {
                Application.OpenForms[Application.OpenForms.Count - 1].Close();
            }
            legend.ShowDialog();
        }

        private void legendButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.legendButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_table_legend3));
        }

        private void legendButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.legendButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_table_legend2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.legendButtonPictureBox, "Show the table legend.");
        }

        private void legendButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.legendButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_table_legend));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.legendButtonPictureBox, "Show the table legend.");
        }

        private void legendButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.legendButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_table_legend));
        }

        private void saveAsCsvButtonPictureBox_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            saveFileDialog1.Filter = "CSV|*.csv";
            if (this.saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // create one file gridview.csv in writing mode using streamwriter
                StreamWriter sw = new StreamWriter("c:\\gridview.csv");
                // now add the gridview header in csv file suffix with "," delimeter except last one
                for (int i = 0; i < importedfileDataGridView.Columns.Count; i++)
                {
                    sw.Write(importedfileDataGridView.Columns[i].HeaderText);
                    if (i != importedfileDataGridView.Columns.Count)
                    {
                        sw.Write(",");
                    }
                }
                // add new line
                sw.Write(sw.NewLine);
                // iterate through all the rows within the gridview
                foreach (DataGridViewRow dr in importedfileDataGridView.Rows)
                {
                    // iterate through all colums of specific row
                    for (int i = 0; i < importedfileDataGridView.Columns.Count; i++)
                    {
                        // write particular cell to csv file
                        sw.Write(dr.Cells[i]);
                        if (i != importedfileDataGridView.Columns.Count)
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
            progressBar1.MarqueeAnimationSpeed = 0;
        }

        private void saveAsCsvButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.saveAsCsvButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_csv_save3));
        }

        private void saveAsCsvButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.saveAsCsvButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_csv_save2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.saveAsCsvButtonPictureBox, "Save as a CSV file.");
        }

        private void saveAsCsvButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.saveAsCsvButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_csv_save));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.saveAsCsvButtonPictureBox, "Save as a CSV file.");
        }

        private void saveAsCsvButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.saveAsCsvButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_csv_save));
        }

        private void saveAsXmlButtonPictureBox_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            saveFileDialog1.Filter = "XML|*.xml";
            if (this.saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                DataTable dt = (DataTable)this.importedfileDataGridView.DataSource;
                dt.WriteXml(this.saveFileDialog1.FileName, XmlWriteMode.WriteSchema);
            }
            progressBar1.MarqueeAnimationSpeed = 0;
        }

        private void saveAsXmlButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.saveAsXmlButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xml_save3));
        }

        private void saveAsXmlButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.saveAsXmlButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xml_save2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.saveAsXmlButtonPictureBox, "Save as an XML file.");
        }

        private void saveAsXmlButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.saveAsXmlButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xml_save));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.saveAsXmlButtonPictureBox, "Save as an XML file.");
        }

        private void saveAsXmlButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.saveAsXmlButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xml_save));
        }

        private void benchmarkGoPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.benchmarkGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go3));
        }

        private void benchmarkGoPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.benchmarkGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.benchmarkGoPictureBox, "Run the tool!");
        }

        private void benchmarkGoPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.benchmarkGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.benchmarkGoPictureBox, "Run the tool!");
        }

        private void benchmarkGoPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.benchmarkGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
        }

        private void selectAllCellLengthCheckerPictureBox_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < cellLengthCheckerListBox.Items.Count; i++)
            {
                cellLengthCheckerListBox.SetSelected(i, true);
            }
        }

        private void clearAllCellLengthCheckerPictureBox_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < cellLengthCheckerListBox.Items.Count; i++)
            {
                cellLengthCheckerListBox.SetSelected(i, false);
            }
        }

        private void cellLengthCheckerGoButtonPictureBox_Click(object sender, EventArgs e)
        {
            int a = 0;
            String reqItem;
            if (checkToolsMaxLengthTextBox.Text.Length == 0)
            {
                MessageBox.Show("You did not enter a length!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            int length = int.Parse(checkToolsMaxLengthTextBox.Text);
            importFormatProgressBar.Value = 50;
            foreach (Object selecteditem in cellLengthCheckerListBox.SelectedItems)
            {
                a++;
                reqItem = selecteditem as String;
                int lengthCharacterCurIndex = cellLengthCheckerListBox.Items.IndexOf(reqItem);
                if (lengthCharacterCurIndex >= 0)
                {

                    for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                    {

                        var value = importedfileDataGridView.Rows[i].Cells[lengthCharacterCurIndex].Value.ToString();
                        //MessageBox.Show("value "+value+"reqitem "+reqItem);
                        if (value.Length > length)
                        {
                            importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[lengthCharacterCurIndex];
                            importFormatProgressBar.Value = 100;
                            MessageBox.Show("The value '" + value + "'" + " in column " + selecteditem + ", line " + (i + 1) + " is too long", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
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
            importFormatProgressBar.Value = 100;
            MessageBox.Show("All columns/rows are under " + length, "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
        }

        private void clearAllNullCheckerButtonPictureBox_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < nullCheckerListBox.Items.Count; i++)
            {
                nullCheckerListBox.SetSelected(i, false);
            }
        }

        private void selectAllNullCheckerButtonPictureBox_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < nullCheckerListBox.Items.Count; i++)
            {
                nullCheckerListBox.SetSelected(i, true);
            }
        }

        private void nullCheckerGoButtonPictureBox_Click(object sender, EventArgs e)
        {
            int a = 0;
            String reqItem;
            importFormatProgressBar.Value = 50;
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
                            importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[nullCheckCurIndex];
                            importFormatProgressBar.Value = 100;
                            MessageBox.Show("NULL value found in column " + "'" + reqItem + "'" + " at line " + (i + 1), "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);

                            return;
                        }
                    }
                }
            }
            if (a == 0)
            {
                importFormatProgressBar.Value = 0;
                MessageBox.Show("You did not select a column!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            importFormatProgressBar.Value = 100;
            MessageBox.Show("no NULL value!", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
        }

        private void clearAllSpecialCharacterCheckerButtonPictureBox_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < specialCharacterCheckerListBox.Items.Count; i++)
            {
                specialCharacterCheckerListBox.SetSelected(i, false);
            }
        }

        private void selectAllSpecialCharacterCheckerButtonPictureBox_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < specialCharacterCheckerListBox.Items.Count; i++)
            {
                specialCharacterCheckerListBox.SetSelected(i, true);
            }
        }

        private void specialCharacterCheckerGoButtonPictureBox_Click(object sender, EventArgs e)
        {
            int a = 0;
            String reqItem;
            String specialChar = checkToolsSpecialCharacterTextBox.Text;
            if (checkToolsSpecialCharacterTextBox.Text.Length == 0)
            {
                MessageBox.Show("You did not enter a special character!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            importFormatProgressBar.Value = 50;
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
                        if (value.Contains(specialChar) == true)
                        {
                            importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[specialCharacterCurIndex];
                            importFormatProgressBar.Value = 100;
                            MessageBox.Show("'" + specialChar + "'" + " WAS found in the column " + "'" + selecteditem + "'" + " at line " + (i + 1), "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);

                            return;
                        }
                    }
                }
            }
            if (a == 0)
            {
                MessageBox.Show("You did not select a column!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                importFormatProgressBar.Value = 0;
                return;
            }
            importFormatProgressBar.Value = 100;
            MessageBox.Show("'" + specialChar + "'" + " WAS NOT FOUND!", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
        }

        private void clearAllDateCheckerButtonPictureBox_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dateCheckerListBox.Items.Count; i++)
            {
                dateCheckerListBox.SetSelected(i, false);
            }
        }

        private void selectAllDateCheckerButtonPictureBox_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dateCheckerListBox.Items.Count; i++)
            {
                dateCheckerListBox.SetSelected(i, true);
            }
        }

        private void dateCheckerGoButtonPictureBox_Click(object sender, EventArgs e)
        {
            int a = 0;
            String reqItem;
            importFormatProgressBar.Value = 50;
            foreach (Object selecteditem in dateCheckerListBox.SelectedItems)
            {
                a++;
                reqItem = selecteditem as String;
                int dateFormatCurIndex = dateCheckerListBox.Items.IndexOf(reqItem);
                if (dateFormatCurIndex >= 0)
                {
                    for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                    {
                        var value = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex].Value.ToString();
                        if (dateCheckerFindNullCheckbox.Checked)
                        {
                            if (value == " " || value == "" || value == null)
                            {
                                importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                importFormatProgressBar.Value = 100;
                                MessageBox.Show("NULL at line " + (i + 1) + "\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   NULL at line " + (i + 1) + "\r\nMake sure that the date is in the format: yyyymmdd");
                                return;
                            }
                        }

                        if (value.Length == 8)
                        {
                            try
                            {

                                int year = int.Parse(value.Substring(0, 4));
                                int month = int.Parse(value.Substring(4, 2));
                                int day = int.Parse(value.Substring(6, 2));

                                if (year > 2200)
                                {
                                    importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                    importFormatProgressBar.Value = 100;
                                    MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyymmdd");
                                    return;
                                }

                                if (month > 12)
                                {
                                    importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                    importFormatProgressBar.Value = 100;
                                    MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyymmdd");
                                    return;
                                }

                                if (month < 01)
                                {
                                    importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                    importFormatProgressBar.Value = 100;
                                    MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyymmdd");
                                    return;
                                }

                                if (day > 31)
                                {
                                    importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                    importFormatProgressBar.Value = 100;
                                    MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyymmdd");
                                    return;
                                }

                                if (day < 01)
                                {
                                    importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                    importFormatProgressBar.Value = 100;
                                    MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyymmdd");
                                    return;
                                }
                            }
                            catch
                            {
                                importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                importFormatProgressBar.Value = 100;
                                MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The value has characters that are not numbers.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The value has characters that are not numbers.\r\nMake sure that the date is in the format: yyyymmdd");
                                return;
                            }
                        }
                        else
                        {
                            if (value.Length > 0)
                            {
                                importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                importFormatProgressBar.Value = 100;
                                MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyymmdd");
                                return;
                            }
                        }
                    }
                }
            }
            if (a == 0)
            {
                importFormatProgressBar.Value = 0;
                MessageBox.Show("You did not select a column!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            MessageBox.Show("Dates are OK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            importFormatProgressBar.Value = 100;
            systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Dates are OK");
            return;
        }

        private void fileSweepUploadFilesPictureBox_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;

            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { ValidateNames = true, Multiselect = true })
                {

                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        fileSweepDataGridView.Columns.Clear();
                        fileSweepDataGridView.Rows.Clear();
                        fileSweepDataGridView.Columns.Add("FileName", "File Name");

                        DataGridViewColumn columnWidth0 = fileSweepDataGridView.Columns[0];
                        columnWidth0.Width = 200;

                        foreach (String file in ofd.SafeFileNames)
                        {
                            fileSweepDataGridView.Rows.Add(file);
                        }
                        serverSelect7.Enabled = true;
                        fileSweepDatabaseComboBox.Enabled = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.MarqueeAnimationSpeed = 0;
        }

        private void btnUp_Click(object sender, EventArgs e)
        {
            DataGridView dgv = fileSweepDataGridView;
            try
            {
                int totalRows = dgv.Rows.Count;
                // get index of the row for the selected cell
                int rowIndex = dgv.SelectedCells[0].OwningRow.Index;
                if (rowIndex == 0)
                    return;
                // get index of the column for the selected cell
                int colIndex = dgv.SelectedCells[0].OwningColumn.Index;
                DataGridViewRow selectedRow = dgv.Rows[rowIndex];
                dgv.Rows.Remove(selectedRow);
                dgv.Rows.Insert(rowIndex - 1, selectedRow);
                dgv.ClearSelection();
                dgv.Rows[rowIndex - 1].Cells[colIndex].Selected = true;
            }
            catch { }
        }

        private void btnDown_Click(object sender, EventArgs e)
        {
            DataGridView dgv = fileSweepDataGridView;
            try
            {
                int totalRows = dgv.Rows.Count;
                // get index of the row for the selected cell
                int rowIndex = dgv.SelectedCells[0].OwningRow.Index;
                if (rowIndex == totalRows - 1)
                    return;
                // get index of the column for the selected cell
                int colIndex = dgv.SelectedCells[0].OwningColumn.Index;
                DataGridViewRow selectedRow = dgv.Rows[rowIndex];
                dgv.Rows.Remove(selectedRow);
                dgv.Rows.Insert(rowIndex + 1, selectedRow);
                dgv.ClearSelection();
                dgv.Rows[rowIndex + 1].Cells[colIndex].Selected = true;
            }
            catch { }
        }

        private void serverSelect7_SelectedIndexChanged(object sender, EventArgs e)
        {
            fileSweepProgressBar.Value = 0;
            progressBar1.MarqueeAnimationSpeed = 1;
            fileSweepProgressBar.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect7.Text + "; Initial Catalog = master; Integrated Security = True");
            try
            {
                conn.Open();
                SqlCommand sc = new SqlCommand("SELECT name FROM [master].[sys].[databases] where name <> 'master' and name <> 'tempdb' and name <> 'model' and name <> 'msdb' and name <> 'DBAtools'", conn);
                SqlDataReader reader;
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                fileSweepDatabaseComboBox.DataSource = dt;
                fileSweepDatabaseComboBox.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading SQL server: " + serverSelect7.Text + "...Done.");
            }
            catch
            {
                conn.Close();
                MessageBox.Show("Unable to connect to the server. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                fileSweepProgressBar.Value = 0;
                return;
            }
            progressBar1.MarqueeAnimationSpeed = 0;
            fileSweepProgressBar.Value = 100;
        }

        private void fileSweepDatabaseComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                fileSweepDataGridView.Columns.Remove("FileSweep");
            }
            catch { }
            
            progressBar1.MarqueeAnimationSpeed = 1;
            benchmarkProgressBar.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect7.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            SqlCommand sc = new SqlCommand("use " + fileSweepDatabaseComboBox.Text + " SELECT filesweepid as name FROM filesweep order by name", conn);
            SqlDataReader reader;
            SqlCommand sc1 = new SqlCommand("use " + databaseSelect4.Text + " select distinct timefrom as name from RunList rl inner join rundet rd on rd.runlistno=rl.runlistno where rd.ItemName='PayoutTypeNo' and rd.ItemValue=(select payouttypeno from PayoutType where payouttypeid='" + payoutTypeSelect.Text + "') and rl.rectype='pay' and DatFrom='" + payoutSelect.Text + "' and rl.finalizestatus='p' order by 1 desc", conn);

            try
            {
                fileSweepGoPictureBox.Visible = true;
                fileSweepGoPictureBox.Enabled = true;
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                DataGridViewComboBoxColumn col = new DataGridViewComboBoxColumn();
                fileSweepDataGridView.Columns.Add(col);
                col.DataSource = dt;
                for (int i = 0; i < fileSweepDataGridView.RowCount; i++)
                {
                    fileSweepDataGridView.Rows[i].Cells[1].Value = null;
                    DataGridViewComboBoxCell c = new DataGridViewComboBoxCell();
                    c.DataSource = dt;
                    c.DisplayMember = "name";
                    fileSweepDataGridView.Rows[i].Cells[1] = c;
                    fileSweepDataGridView.Columns[1].Name = "FileSweep";
                    fileSweepDataGridView.Columns[1].HeaderText = "File Sweep";
                    DataGridViewColumn columnWidth1 = fileSweepDataGridView.Columns[1];
                    columnWidth1.Width = 200;
                    
                }
                conn.Close();
                connectionStatus.Visible = true;
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading database: " + fileSweepDatabaseComboBox.Text + "...Done.");
            }
            catch
            {
                conn.Close();
                progressBar1.MarqueeAnimationSpeed = 0;
                fileSweepProgressBar.Value = 0;
                return;
            }
            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            fileSweepProgressBar.Value = 100;
        }

        //------------------SQL LOADER START------------------------------------------------------

        private void serverSelect_SelectedIndexChanged(object sender, EventArgs e)
        {

            importFormatProgressBar.Value = 0;
            progressBar1.MarqueeAnimationSpeed = 1;
            importFormatProgressBar.Value = 20;
            importFormatProgressBar.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
            try
            {
                conn.Open();
                SqlCommand sc = new SqlCommand("SELECT name FROM [master].[sys].[databases] where name <> 'master' and name <> 'tempdb' and name <> 'model' and name <> 'msdb' and name <> 'DBAtools'", conn);
                SqlDataReader reader;
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                databaseSelect.DataSource = dt;
                databaseSelect.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading SQL server: " + serverSelect.Text + "...Done.");
            }
            catch
            {
                conn.Close();
                MessageBox.Show("Unable to connect to the server. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                importFormatProgressBar.Value = 0;
                return;
            }
            progressBar1.MarqueeAnimationSpeed = 0;
            importFormatProgressBar.Value = 100;
        }

        private void serverSelect2_SelectedIndexChanged(object sender, EventArgs e)
        {

            sqlQueryProgressBar.Value = 0;
            progressBar1.MarqueeAnimationSpeed = 1;
            sqlQueryProgressBar.Value = 20;
            sqlQueryProgressBar.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect2.Text + "; Initial Catalog = master; Integrated Security = True");
            try
            {
                conn.Open();
                SqlCommand sc = new SqlCommand("SELECT name FROM [master].[sys].[databases] where name <> 'master' and name <> 'tempdb' and name <> 'model' and name <> 'msdb' and name <> 'DBAtools'", conn);
                SqlDataReader reader;
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                databaseSelect2.DataSource = dt;
                databaseSelect2.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading SQL server: " + serverSelect.Text + "...Done.");
            }
            catch
            {
                conn.Close();
                MessageBox.Show("Unable to connect to the server. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                sqlQueryProgressBar.Value = 0;
                return;
            }
            progressBar1.MarqueeAnimationSpeed = 0;
            sqlQueryProgressBar.Value = 100;
        }



        private void serverSelect4_SelectedIndexChanged(object sender, EventArgs e)
        {

            importFormatProgressBar.Value = 0;
            progressBar1.MarqueeAnimationSpeed = 1;
            benchmarkProgressBar.Value = 20;
            benchmarkProgressBar.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect4.Text + "; Initial Catalog = master; Integrated Security = True");
            try
            {
                conn.Open();
                SqlCommand sc = new SqlCommand("SELECT name FROM [master].[sys].[databases] where name <> 'master' and name <> 'tempdb' and name <> 'model' and name <> 'msdb' and name <> 'DBAtools'", conn);
                SqlDataReader reader;
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                databaseSelect4.DataSource = dt;
                databaseSelect4.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading SQL server: " + serverSelect4.Text + "...Done.");
                benchmarkProgressBar.Value = 100;
            }
            catch
            {
                conn.Close();
                MessageBox.Show("Unable to connect to the server. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                benchmarkProgressBar.Value = 0;
            }
            progressBar1.MarqueeAnimationSpeed = 0;
        }


        private void serverSelect5_SelectedIndexChanged(object sender, EventArgs e)
        {

            apiReadinessProgressBar.Value = 0;
            progressBar1.MarqueeAnimationSpeed = 1;
            apiReadinessProgressBar.Value = 20;
            apiReadinessProgressBar.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect5.Text + "; Initial Catalog = master; Integrated Security = True");
            try
            {
                conn.Open();
                SqlCommand sc = new SqlCommand("SELECT name FROM [master].[sys].[databases] where name <> 'master' and name <> 'tempdb' and name <> 'model' and name <> 'msdb' and name <> 'DBAtools'", conn);
                SqlDataReader reader;
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                databaseSelect5.DataSource = dt;
                databaseSelect5.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading SQL server: " + serverSelect5.Text + "...Done.");
            }
            catch
            {
                conn.Close();
                MessageBox.Show("Unable to connect to the server. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                apiReadinessProgressBar.Value = 0;
                return;
            }
            progressBar1.MarqueeAnimationSpeed = 0;
            apiReadinessProgressBar.Value = 100;
        }
        private void serverSelect6_SelectedIndexChanged(object sender, EventArgs e)
        {
            envChangesProgressBar.Value = 0;
            progressBar1.MarqueeAnimationSpeed = 1;
            envChangesProgressBar.Value = 20;
            envChangesProgressBar.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect6.Text + "; Initial Catalog = master; Integrated Security = True");
            try
            {
                conn.Open();
                SqlCommand sc = new SqlCommand("SELECT name FROM [master].[sys].[databases] where name <> 'master' and name <> 'tempdb' and name <> 'model' and name <> 'msdb' and name <> 'DBAtools'", conn);
                SqlDataReader reader;
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                databaseSelect6.DataSource = dt;
                databaseSelect6.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading SQL server: " + serverSelect6.Text + "...Done.");
            }
            catch
            {
                conn.Close();
                MessageBox.Show("Unable to connect to the server. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                envChangesProgressBar.Value = 0;
                return;
            }
            progressBar1.MarqueeAnimationSpeed = 0;
            envChangesProgressBar.Value = 100;
        }

        private void runquery_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            System.Threading.Thread.Sleep(25);
            sqlQueryProgressBar.Value = 20;
            sqlQueryProgressBar.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect2.Text + "; Initial Catalog = master; Integrated Security = True");

            try
            {
                string ID = databaseSelect2.SelectedValue.ToString();
                conn.Open();
                var select = "USE " + databaseSelect2.Text + " " + queryWindow.Text;
                if (queryWindow.Text.Equals("select * from tranhis", StringComparison.InvariantCultureIgnoreCase))
                {
                    DialogResult result = MessageBox.Show("Performing a SELECT * FROM TRANHIS is insane. Continue?", "Data Analysis Tool", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (result == DialogResult.No)
                    {
                        progressBar1.MarqueeAnimationSpeed = 0;
                        sqlQueryProgressBar.Value = 0;
                        return;
                    }
                }
                var conn2 = new SqlConnection(@"Data Source = " + serverSelect2.Text + "; Initial Catalog = master; Integrated Security = True");
                var dataAdapter = new SqlDataAdapter(select, conn2);
                var commandBuilder = new SqlCommandBuilder(dataAdapter);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                sqlQueryDataGridView.ReadOnly = true;
                sqlQueryDataGridView.DataSource = ds.Tables[0];
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                conn.Close();
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Running query against: " + databaseSelect2.Text + "...Done.");
                sqlQueryProgressBar.Value = 100;
            }
            catch
            {
                conn.Close();
                MessageBox.Show("Unable to run query. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                sqlQueryProgressBar.Value = 0;
                return;
            }
            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            importFormatProgressBar.Value = 100;
        }

        private void databaseSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            importFormatProgressBar.Value = 20;
            importFormatProgressBar.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            SqlCommand sc = new SqlCommand("use " + databaseSelect.Text + " SELECT table_name AS name FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' order by TABLE_NAME", conn);
            SqlCommand scVersion = new SqlCommand("use " + databaseSelect.Text + " SELECT codetype FROM entityfield", conn);
            SqlDataReader reader;

            try
            {
                reader = scVersion.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                icmVersion.Visible = true;
                icmVersion.Text = "v.7.0";
            }
            catch
            {
                icmVersion.Visible = true;
                icmVersion.Text = "v.2018";
            }

            try
            {
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                tableSelect.DataSource = dt;
                tableSelect.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading database: " + databaseSelect.Text + "...Done.");
                seperator3ToolStripStatusLabel.Visible = true;
                sqlRowCountToolStripStatusLabel.Visible = true;
                sqlCounterToolStripStatusLabel.Visible = true;
            }
            catch
            {
                return;
            }

            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            importFormatProgressBar.Value = 100;
        }

        //databaseSelect2 not used right now
        //databaseSelect3 not used right now

        private void databaseSelect4_SelectedIndexChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("should not hit this");
            //payoutSelect.SelectedIndex = -1;
            //payoutTypeSelect.SelectedIndex = -1;
            progressBar1.MarqueeAnimationSpeed = 1;
            benchmarkProgressBar.Value = 20;
            benchmarkProgressBar.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect4.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            SqlCommand sc = new SqlCommand("use " + databaseSelect4.Text + " SELECT payouttypeid as name FROM payouttype  order by name", conn);
            SqlDataReader reader;

            try
            {
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                payoutTypeSelect.DataSource = dt;
                payoutTypeSelect.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading database: " + databaseSelect.Text + "...Done.");
                seperator3ToolStripStatusLabel.Visible = true;
                sqlRowCountToolStripStatusLabel.Visible = true;
                sqlCounterToolStripStatusLabel.Visible = true;
            }
            catch
            {
                conn.Close();
                progressBar1.MarqueeAnimationSpeed = 0;
                benchmarkProgressBar.Value = 0;
                return;
            }
            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            benchmarkProgressBar.Value = 100;
        }


        private void payoutTypeSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            benchmarkProgressBar.Value = 20;
            benchmarkProgressBar.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect4.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            //data select
            SqlDataReader reader;
            SqlCommand sc1 = new SqlCommand("use " + databaseSelect4.Text + " select distinct datfrom as name from RunList rl inner join rundet rd on rd.runlistno=rl.runlistno where rd.ItemName='PayoutTypeNo' and rd.ItemValue=(select payouttypeno from PayoutType where payouttypeid='" + payoutTypeSelect.Text + "') and rl.rectype='pay' and rl.finalizestatus='p' order by 1 desc", conn);
            SqlCommand sc2 = new SqlCommand("use " + databaseSelect4.Text + " select distinct datfrom as name from RunList rl inner join rundet rd on rd.runlistno=rl.runlistno where rd.ItemName='PayoutTypeNo' and rd.ItemValue=(select payouttypeno from PayoutType where payouttypeid='" + payoutTypeSelect.Text + "') and rl.rectype='pay' and rl.finalizestatus='f' order by 1 desc", conn);
            SqlCommand sc3 = new SqlCommand("use " + databaseSelect4.Text + " select distinct datfrom as name from RunList rl inner join rundet rd on rd.runlistno=rl.runlistno where rd.ItemName='PayoutTypeNo' and rd.ItemValue=(select payouttypeno from PayoutType where payouttypeid='" + payoutTypeSelect.Text + "') and rl.rectype='pay' and rl.finalizestatus='r' order by 1 desc", conn);


            try
            {
                if (pendingRadioButton.Checked == true)
                {
                    reader = sc1.ExecuteReader();
                }
                else if (finalizedRadioButton.Checked == true)
                {
                    reader = sc2.ExecuteReader();
                }
                else if (reversedRadioButton.Checked == true)
                {
                    reader = sc3.ExecuteReader();
                }
                else
                {
                    return;
                }
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                payoutSelect.DataSource = dt;
                payoutSelect.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading payouts: " + payoutTypeSelect.Text + "...Done.");
                seperator3ToolStripStatusLabel.Visible = true;
                sqlRowCountToolStripStatusLabel.Visible = true;
                sqlCounterToolStripStatusLabel.Visible = true;
            }
            catch
            {
                conn.Close();
                progressBar1.MarqueeAnimationSpeed = 0;
                benchmarkProgressBar.Value = 0;
                return;
            }
            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            benchmarkProgressBar.Value = 100;
        }

        private void payoutSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            benchmarkProgressBar.Value = 20;
            benchmarkProgressBar.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect4.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            //data select
            SqlDataReader reader;
            SqlCommand sc1 = new SqlCommand("use " + databaseSelect4.Text + " select distinct timefrom as name from RunList rl inner join rundet rd on rd.runlistno=rl.runlistno where rd.ItemName='PayoutTypeNo' and rd.ItemValue=(select payouttypeno from PayoutType where payouttypeid='" + payoutTypeSelect.Text + "') and rl.rectype='pay' and DatFrom='" + payoutSelect.Text + "' and rl.finalizestatus='p' order by 1 desc", conn);
            SqlCommand sc2 = new SqlCommand("use " + databaseSelect4.Text + " select distinct timefrom as name from RunList rl inner join rundet rd on rd.runlistno=rl.runlistno where rd.ItemName='PayoutTypeNo' and rd.ItemValue=(select payouttypeno from PayoutType where payouttypeid='" + payoutTypeSelect.Text + "') and rl.rectype='pay' and DatFrom='" + payoutSelect.Text + "' and rl.finalizestatus='f' order by 1 desc", conn);
            SqlCommand sc3 = new SqlCommand("use " + databaseSelect4.Text + " select distinct timefrom as name from RunList rl inner join rundet rd on rd.runlistno=rl.runlistno where rd.ItemName='PayoutTypeNo' and rd.ItemValue=(select payouttypeno from PayoutType where payouttypeid='" + payoutTypeSelect.Text + "') and rl.rectype='pay' and DatFrom='" + payoutSelect.Text + "' and rl.finalizestatus='r' order by 1 desc", conn);


            try
            {
                if (pendingRadioButton.Checked == true)
                {
                    reader = sc1.ExecuteReader();
                }
                else if (finalizedRadioButton.Checked == true)
                {
                    reader = sc2.ExecuteReader();
                }
                else if (reversedRadioButton.Checked == true)
                {
                    reader = sc3.ExecuteReader();
                }
                else
                {
                    return;
                }
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                payoutTimeSelect.DataSource = dt;
                payoutTimeSelect.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading payouts: " + payoutTypeSelect.Text + "...Done.");
                seperator3ToolStripStatusLabel.Visible = true;
                sqlRowCountToolStripStatusLabel.Visible = true;
                sqlCounterToolStripStatusLabel.Visible = true;
            }
            catch
            {
                conn.Close();
                progressBar1.MarqueeAnimationSpeed = 0;
                benchmarkProgressBar.Value = 0;
                return;
            }
            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            benchmarkProgressBar.Value = 100;
        }

        private void tableSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            importFormatProgressBar.Value = 20;
            importFormatProgressBar.Value = 40;
            string ID = databaseSelect.SelectedValue.ToString();
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            SqlCommand sc;
            if (importFormatShowOpenImportFormatsButton.Checked == true)
            {
                sc = new SqlCommand("use " + databaseSelect.Text + " select importformatid as name from ImportFormat", conn);
            }
            else
            {
                sc = new SqlCommand("use " + databaseSelect.Text + " select importformatid as name from ImportFormat where prosta=1", conn);
            }

            SqlDataReader reader;
            try
            {
                var select = "USE " + databaseSelect.Text + " SELECT top 20000 * FROM " + tableSelect.Text;
                var conn2 = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
                var dataAdapter = new SqlDataAdapter(select, conn2);
                var commandBuilder = new SqlCommandBuilder(dataAdapter);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                dataGridView2.ReadOnly = true;
                dataGridView2.DataSource = ds.Tables[0];
                sqlCounterToolStripStatusLabel.Text = dataGridView2.Rows.Count.ToString();

                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                ifSelect.DataSource = dt;
                ifSelect.DisplayMember = "name";
                conn.Close();
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading table: " + tableSelect.Text + "...Done.");
            }
            catch
            {
                return;
            }

            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            importFormatProgressBar.Value = 100;
        }



        private void ifSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            importFormatProgressBar.Value = 20;
            importFormatProgressBar.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            try
            {
                var select = "USE " + databaseSelect.Text + " SELECT IMF.ImportFormatId,IMF.Delimiter,IMF.HeaderRows,IMF.RecType,IMFE.InEntName,IMFF.ImportFormatFieldId,IMFF.FieldSeq,IMFF.FieldLength,IMFF.IgnoreField, ef.* FROM ImportFormat IMF INNER JOIN ImportFormatEntity IMFE ON IMF.ImportFormatNo= IMFE.ImportFormatNo INNER JOIN ImportFormatField IMFF ON IMF.ImportFormatNo = IMFF.ImportFormatNo  left JOIN EntityField EF ON ef.entname=imfe.inentname and ef.fldname=IMFF.ImportFormatFieldId where imf.importformatid = " + @"'" + ifSelect.Text + @"'" + "  and IMF.QBQueryNo is null order by imff.FieldSeq";
                var select2 = "USE " + databaseSelect.Text + " SELECT IMFF.ImportFormatFieldId FROM ImportFormat IMF INNER JOIN ImportFormatEntity IMFE ON IMF.ImportFormatNo= IMFE.ImportFormatNo INNER JOIN ImportFormatField IMFF ON IMF.ImportFormatNo = IMFF.ImportFormatNo  left JOIN EntityField EF ON ef.entname=imfe.inentname and ef.fldname=IMFF.ImportFormatFieldId where imf.importformatid = " + @"'" + ifSelect.Text + @"'" + "  and IMF.QBQueryNo is null order by imff.FieldSeq";
                var conn2 = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
                var dataAdapter = new SqlDataAdapter(select, conn2);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                importformatDataGridView.DataSource = ds.Tables[0];

                var iffidArray2 = importformatDataGridView.Rows.Cast<DataGridViewRow>()
                    .Select(x => x.Cells[5].Value.ToString().Trim()).ToArray();

                //gives me the InEntName of the import format
                var selectInEntName = "USE " + databaseSelect.Text + " select top 1 ife.InEntName from ImportFormat i inner join importformatentity ife on i.ImportFormatNo=ife.ImportFormatNo left join ImportFormatFieldMapping iffm on iffm.ImportFormatEntityNo=ife.ImportFormatEntityNo where i.ImportFormatId=" + @"'" + ifSelect.Text + @"'";
                var dataAdapter8 = new SqlDataAdapter(selectInEntName, conn);
                var ds8 = new DataSet();
                dataAdapter8.Fill(ds8);
                stagedDataGridView.DataSource = ds8.Tables[0];
                var inEntName = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                var inEntNameVar = stagedDataGridView.Rows[0].Cells[0].Value.ToString();

                if (inEntNameVar == "InAddress")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InAdjustmentHis")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InAssignment")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBroker")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerAdj")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerContract")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerCustomer")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerDetail")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerHierarchy")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerHold")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerLicense")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerReserveHis")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerRoleBroker")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerVendor")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCarrier")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCertificate")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCertificateDet")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCmsMarx")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCmsMmr")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCmsTrr")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCodSet")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCustomer")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCustomerApplication")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCustomerMatch")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCustPolicy")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InEducation")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InEntityRef")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InExtCrossRef")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InFile")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InFileImportFile")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InFileImportParm")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InFileImportRequest")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InFileRunList")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InIdentSet")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InMatchRule")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InPerfHis")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InPrepayBalanceAdjustment")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProAppointment")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProAppointmentDet")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProBackground")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProContract")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProContractDet")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProducer")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProducts")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProductsLicense")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProInsurance")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProLicense")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProLicenseDet")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InterestDetail")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InterestSet")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InTimeSheet")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InTranDefault")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InTranHead")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InVendor")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InVoucher")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }

                reqListBox.Items.Clear();
                int a = 0;
                for (int i = 0; i < iffidArray2.Length; i++)
                {
                    a++;
                    reqListBox.Items.Add(a + ". " + iffidArray2[i].ToString());
                }

                dateListBox.Items.Clear();

                a = 0;
                for (int i = 0; i < iffidArray2.Length; i++)
                {
                    a++;
                    dateListBox.Items.Add(a + ". " + iffidArray2[i].ToString());
                }
                conn.Close();

                sqlCounterToolStripStatusLabel.Text = dataGridView2.Rows.Count.ToString();
                importFormatRowCountToolStripStatusLabel.Text = importformatDataGridView.Rows.Count.ToString();
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading import format: " + ifSelect.Text + "...Done.");
                seperator2ToolStripStatusLabel.Visible = true;
                ifRowCountToolStripStatusLabel.Visible = true;
                importFormatRowCountToolStripStatusLabel.Visible = true;
            }
            catch
            {
                return;
            }

            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            importFormatProgressBar.Value = 100;
        }

        //------------------SQL LOADER END------------------------------------------------------

        private void groupByErrorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            importFormatProgressBar.Value = 0;
            importFormatProgressBar.Value = 10;

            //global vars
            progressBar1.MarqueeAnimationSpeed = 1;
            var ifCount = "USE " + databaseSelect.Text + " SELECT IMFF.FieldSeq FROM ImportFormat IMF INNER JOIN ImportFormatEntity IMFE ON IMF.ImportFormatNo= IMFE.ImportFormatNo INNER JOIN ImportFormatField IMFF ON IMF.ImportFormatNo = IMFF.ImportFormatNo where imf.importformatid = " + @"'" + ifSelect.Text + @"'" + "  and IMF.QBQueryNo is null order by imff.FieldSeq";


            if (importedfileDataGridView.Rows.Count == 0)

            {
                MessageBox.Show("No file imported. \nPlease open a file.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                importFormatProgressBar.Value = 0;
                return;
            }

            if (ifSelect.Text == "")

            {
                DialogResult result = MessageBox.Show("No IF selected. \nPlease make sure you are connected to ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                importFormatProgressBar.Value = 0;
                return;
            }

            if (databaseSelect.Text != "")
            {

                DialogResult result2 = MessageBox.Show("The DAT will check against the " + ifSelect.Text + " Import Format.\nContinue?", "Data Analysis Tool", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                if (result2 == DialogResult.No)
                {
                    progressBar1.MarqueeAnimationSpeed = 0;
                    importFormatProgressBar.Value = 0;
                    return;
                }
            }

            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            SqlCommand sc = new SqlCommand("use " + databaseSelect.Text + " select importformatid as name from ImportFormat", conn);

            //for version 7.0
            var selectCodeType1 = "USE " + databaseSelect.Text + " SELECT ef.codetype FROM ImportFormat IMF INNER JOIN ImportFormatEntity IMFE ON IMF.ImportFormatNo= IMFE.ImportFormatNo INNER JOIN ImportFormatField IMFF ON IMF.ImportFormatNo = IMFF.ImportFormatNo  left JOIN EntityField EF ON ef.entname=imfe.inentname and ef.fldname=IMFF.ImportFormatFieldId where imf.importformatid = " + @"'" + ifSelect.Text + @"'" + "  and IMF.QBQueryNo is null and ef.valuetype=1 order by imff.FieldSeq";
            //for version 2018
            var selectCodeType2 = "USE " + databaseSelect.Text + " SELECT ct.codetypeid FROM ImportFormat IMF INNER JOIN ImportFormatEntity IMFE ON IMF.ImportFormatNo= IMFE.ImportFormatNo INNER JOIN ImportFormatField IMFF ON IMF.ImportFormatNo = IMFF.ImportFormatNo  left JOIN EntityField EF ON ef.entname=imfe.inentname and ef.fldname=IMFF.ImportFormatFieldId left join codetype ct on ef.codetypeno=ct.codetypeno where imf.importformatid = " + @"'" + ifSelect.Text + @"'" + "  and IMF.QBQueryNo is null and ef.valuetype=1 order by imff.FieldSeq";

            var dataAdapter1 = new SqlDataAdapter(selectCodeType1, conn);
            var dataAdapter22 = new SqlDataAdapter(selectCodeType2, conn);
            var ds = new DataSet();
            if (icmVersion.Text == "v.7.0")
            {
                dataAdapter1.Fill(ds);
            }
            if (icmVersion.Text == "v.2018")
            {
                dataAdapter22.Fill(ds);
            }
            else
            {
                dataAdapter22.Fill(ds);
            }

            stagedDataGridView.DataSource = ds.Tables[0];
            var codeArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                    .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

            var selectFieldSeq = "USE " + databaseSelect.Text + " SELECT IMFF.FieldSeq FROM ImportFormat IMF INNER JOIN ImportFormatEntity IMFE ON IMF.ImportFormatNo= IMFE.ImportFormatNo INNER JOIN ImportFormatField IMFF ON IMF.ImportFormatNo = IMFF.ImportFormatNo  left JOIN EntityField EF ON ef.entname=imfe.inentname and ef.fldname=IMFF.ImportFormatFieldId where imf.importformatid = " + @"'" + ifSelect.Text + @"'" + "  and IMF.QBQueryNo is null and ef.valuetype=1 order by imff.FieldSeq";
            var dataAdapter3 = new SqlDataAdapter(selectFieldSeq, conn);
            var ds3 = new DataSet();
            dataAdapter3.Fill(ds3);
            stagedDataGridView.DataSource = ds3.Tables[0];
            var fieldsThatAreCodesArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                    .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

            var selectMaxLength = "USE " + databaseSelect.Text + " SELECT ef.FldName FROM ImportFormat IMF INNER JOIN ImportFormatEntity IMFE ON IMF.ImportFormatNo= IMFE.ImportFormatNo INNER JOIN ImportFormatField IMFF ON IMF.ImportFormatNo = IMFF.ImportFormatNo  left JOIN EntityField EF ON ef.entname=imfe.inentname and ef.fldname=IMFF.ImportFormatFieldId where imf.importformatid = " + @"'" + ifSelect.Text + @"'" + "  and IMF.QBQueryNo is null and ef.MaxLength is not null order by imff.FieldSeq";
            var dataAdapter4 = new SqlDataAdapter(selectMaxLength, conn);
            var ds4 = new DataSet();
            dataAdapter4.Fill(ds4);
            stagedDataGridView.DataSource = ds4.Tables[0];
            var maxLengthFieldArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                    .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

            var selectMaxLengthColumnNumber = "USE " + databaseSelect.Text + " SELECT IMFF.FieldSeq FROM ImportFormat IMF INNER JOIN ImportFormatEntity IMFE ON IMF.ImportFormatNo= IMFE.ImportFormatNo INNER JOIN ImportFormatField IMFF ON IMF.ImportFormatNo = IMFF.ImportFormatNo  left JOIN EntityField EF ON ef.entname=imfe.inentname and ef.fldname=IMFF.ImportFormatFieldId where imf.importformatid = " + @"'" + ifSelect.Text + @"'" + "  and IMF.QBQueryNo is null and ef.MaxLength is not null order by imff.FieldSeq";
            var dataAdapter6 = new SqlDataAdapter(selectMaxLengthColumnNumber, conn);
            var ds6 = new DataSet();
            dataAdapter6.Fill(ds6);
            stagedDataGridView.DataSource = ds6.Tables[0];
            var maxLengthFieldColumnNumberArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                    .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

            var selectMaxLengthValue = "USE " + databaseSelect.Text + " SELECT ef.maxlength FROM ImportFormat IMF INNER JOIN ImportFormatEntity IMFE ON IMF.ImportFormatNo= IMFE.ImportFormatNo INNER JOIN ImportFormatField IMFF ON IMF.ImportFormatNo = IMFF.ImportFormatNo  left JOIN EntityField EF ON ef.entname=imfe.inentname and ef.fldname=IMFF.ImportFormatFieldId where imf.importformatid = " + @"'" + ifSelect.Text + @"'" + "  and IMF.QBQueryNo is null and ef.MaxLength is not null order by imff.FieldSeq";
            var dataAdapter5 = new SqlDataAdapter(selectMaxLengthValue, conn);
            var ds5 = new DataSet();
            dataAdapter5.Fill(ds5);
            stagedDataGridView.DataSource = ds5.Tables[0];
            var maxLengthFieldArrayValue = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                    .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

            //gives me the client name of the selected database
            var selectClientName = "USE " + databaseSelect.Text + " select optval from optset where OptName='ui.title.prefix'";
            var dataAdapter7 = new SqlDataAdapter(selectClientName, conn);
            var ds7 = new DataSet();
            dataAdapter7.Fill(ds7);
            stagedDataGridView.DataSource = ds7.Tables[0];
            var clientName = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                    .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

            var iffidArray = importformatDataGridView.Rows.Cast<DataGridViewRow>()
                    .Select(x => x.Cells[5].Value.ToString().Trim()).ToArray();

            var seqArray = importformatDataGridView.Rows.Cast<DataGridViewRow>()
                .Select(x => x.Cells[6].Value.ToString().Trim()).ToArray();


            int[] fieldsThatAreCodesArrayColumnCount = Array.ConvertAll(fieldsThatAreCodesArray, s => int.Parse(s));

            ArrayList codeValueArray = new ArrayList();
            //this foreach gets the values for all of the codes
            if (icmVersion.Text == "v.7.0")
            {
                foreach (var s in codeArray)
                {
                    var select2 = "USE " + databaseSelect.Text + "  select recval from codset where rectype=" + "'" + s + "'";
                    var dataAdapter2 = new SqlDataAdapter(select2, conn);
                    var ds2 = new DataSet();
                    dataAdapter2.Fill(ds2);
                    stagedDataGridView.DataSource = ds2.Tables[0];

                    foreach (DataGridViewRow dr in stagedDataGridView.Rows)
                    {
                        codeValueArray.Add(dr.Cells[0].Value);
                    }
                }
            }
            else
            {
                foreach (var s in codeArray)
                {
                    var select2 = "USE " + databaseSelect.Text + "   select cv.storedvalue from codevalue cv inner join CodeType ct on ct.CodeTypeNo=cv.CodeTypeNo where ct.CodeTypeId=" + "'" + s + "'";
                    var dataAdapter2 = new SqlDataAdapter(select2, conn);
                    var ds2 = new DataSet();
                    dataAdapter2.Fill(ds2);
                    stagedDataGridView.DataSource = ds2.Tables[0];

                    foreach (DataGridViewRow dr in stagedDataGridView.Rows)
                    {
                        codeValueArray.Add(dr.Cells[0].Value);
                    }
                }
            }
            var intersect = fieldsThatAreCodesArray.Intersect(seqArray);
            int[] intMaxLengthFieldArrayValue = Array.ConvertAll(maxLengthFieldArrayValue, s => int.Parse(s));


            importFormatRowCountToolStripStatusLabel.Text = importformatDataGridView.Rows.Count.ToString();
            sqlCounterToolStripStatusLabel.Text = stagedDataGridView.Rows.Count.ToString();


            {
                System.IO.Directory.CreateDirectory(Application.UserAppDataPath + @"\IF_Error_Files_Data");
                string path = Application.UserAppDataPath + @"\IF_Error_Files_Data\DataAnalysisTool_IFEF_Data_" + DateTime.Now.ToString("MM_dd_yyyy_HHmmss") + ".txt";
                using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
                {
                    using (TextWriter tw = new StreamWriter(fs))
                    {
                        tw.WriteLine("###########################################################################################");
                        tw.WriteLine("########################DataAnalysisTool - Data Used - Import Format#######################");
                        tw.WriteLine("###########################################################################################");
                        tw.WriteLine(DateTime.Now);
                        tw.WriteLine("Server: " + serverSelect.Text);
                        tw.WriteLine("Database: " + databaseSelect.Text);
                        tw.WriteLine("Import Format: " + ifSelect.Text);



                        if (databaseSelect.Text != "")
                        {
                            if (importedfileDataGridView.ColumnCount != importformatDataGridView.RowCount)
                            {
                                tw.WriteLine("This Import Format requires " + importformatDataGridView.RowCount + " columns. You have " + importedfileDataGridView.ColumnCount + ".");
                                tw.WriteLine("This operation has ended. Please correct the column count issue.");
                                tw.WriteLine("EOF.");
                                importFormatProgressBar.Value = 100;
                                MessageBox.Show("Import Format error file has been created. \nLocation: " + path, "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + @">>>   Import Format error file has been created. Location: C:\Program Files (x86)\DataAnalysisTool\Medicare Error Files");
                                progressBar1.MarqueeAnimationSpeed = 0;
                                Process.Start(path);
                                return;
                            }
                            try
                            {
                                foreach (var value in clientName)
                                {
                                    tw.WriteLine("Client: " + value);
                                }
                                tw.WriteLine("");
                                tw.WriteLine("****************************************************");
                                tw.WriteLine("******CONFIGURATION / SYSTEM DATA THAT IS USED******");
                                tw.WriteLine("****************************************************");
                                tw.WriteLine("");

                                tw.WriteLine("---Selected Required Fields---");
                                String reqItem;
                                foreach (Object selecteditem in reqListBox.SelectedItems)
                                {

                                    reqItem = selecteditem as String;
                                    int reqCurIndex = reqListBox.Items.IndexOf(reqItem);
                                    if (reqCurIndex >= 0)
                                    {
                                        tw.WriteLine("Required Column: " + reqItem);
                                    }
                                }
                                tw.WriteLine("---Selected Date Format and Date Columns---");
                                String dateItem;
                                foreach (Object selecteditem in dateListBox.SelectedItems)
                                {

                                    dateItem = selecteditem as String;
                                    int dateCurIndex = dateListBox.Items.IndexOf(dateItem);
                                    if (dateCurIndex >= 0)
                                    {
                                        tw.WriteLine("Date Column: " + dateItem);
                                    }
                                }
                                tw.WriteLine(dateFormat.Text);
                                tw.WriteLine("");
                                tw.WriteLine("");
                                tw.WriteLine("****************************************************");
                                tw.WriteLine("**********SYSTEM DATA PULLED FROM DATABASE**********");
                                tw.WriteLine("****************************************************");
                                tw.WriteLine("");
                                tw.WriteLine("---Predefined Codes in System Configuration---");
                                foreach (var value in codeArray)
                                {
                                    tw.WriteLine("Code: " + value);
                                }
                                foreach (var value in codeValueArray)
                                {
                                    tw.WriteLine("Code Value: " + value);
                                }
                                foreach (int value in fieldsThatAreCodesArrayColumnCount)
                                {
                                    tw.WriteLine("Columns with Codes: " + value);
                                }
                                tw.WriteLine("");
                                tw.WriteLine("---Predefined Field Length Restrictions in System Configuration---");
                                foreach (var value in maxLengthFieldColumnNumberArray)
                                {
                                    tw.WriteLine("Columns with length restrictions: " + value);
                                }

                                foreach (var value in maxLengthFieldArrayValue)
                                {
                                    tw.WriteLine("length restriction: " + value);
                                }
                            }
                            catch { return; }
                        }
                        tw.WriteLine("EOF.");
                    }
                }
            }

            {
                System.IO.Directory.CreateDirectory(Application.UserAppDataPath + @"\IF_Error_Files");
                string path = Application.UserAppDataPath + @"\IF_Error_Files\DataAnalysisTool_IFEF_" + DateTime.Now.ToString("MM_dd_yyyy_HHmmss") + ".txt";
                using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
                {
                    importFormatProgressBar.Value = 20;
                    using (TextWriter tw = new StreamWriter(fs))
                    {
                        tw.WriteLine("###########################################################################################");
                        tw.WriteLine("########################DataAnalysisTool - Import Format Error File########################");
                        tw.WriteLine("###########################################################################################");
                        tw.WriteLine(DateTime.Now);
                        tw.WriteLine("Server: " + serverSelect.Text);
                        tw.WriteLine("Database: " + databaseSelect.Text);
                        tw.WriteLine("Import Format: " + ifSelect.Text);



                        if (databaseSelect.Text != "")
                        {
                            importFormatProgressBar.Value = 30;
                            if (importedfileDataGridView.ColumnCount != importformatDataGridView.RowCount)
                            {
                                tw.WriteLine("This Import Format requires " + importformatDataGridView.RowCount + " columns. You have " + importedfileDataGridView.ColumnCount + ".");
                                tw.WriteLine("This operation has ended. Please correct the column count issue.");
                                MessageBox.Show("Import Format error file has been created. \nLocation: " + path, "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + @">>>   Import Format error file has been created. Location: C:\Program Files (x86)\DataAnalysisTool\Medicare Error Files");
                                progressBar1.MarqueeAnimationSpeed = 0;
                                importFormatProgressBar.Value = 0;
                                Process.Start(path);
                                return;
                            }
                            try
                            {


                                //this foreach gets the values for all of the codes
                                if (icmVersion.Text == "v.7.0")
                                {
                                    foreach (var s in codeArray)
                                    {
                                        var select2 = "USE " + databaseSelect.Text + "  select recval from codset where rectype=" + "'" + s + "'";
                                        var dataAdapter2 = new SqlDataAdapter(select2, conn);
                                        var ds2 = new DataSet();
                                        dataAdapter2.Fill(ds2);
                                        stagedDataGridView.DataSource = ds2.Tables[0];

                                        foreach (DataGridViewRow dr in stagedDataGridView.Rows)
                                        {
                                            codeValueArray.Add(dr.Cells[0].Value);
                                        }
                                    }
                                }
                                else
                                {
                                    foreach (var s in codeArray)
                                    {
                                        var select2 = "USE " + databaseSelect.Text + "   select cv.storedvalue from codevalue cv inner join CodeType ct on ct.CodeTypeNo=cv.CodeTypeNo where ct.CodeTypeId=" + "'" + s + "'";
                                        var dataAdapter2 = new SqlDataAdapter(select2, conn);
                                        var ds2 = new DataSet();
                                        dataAdapter2.Fill(ds2);
                                        stagedDataGridView.DataSource = ds2.Tables[0];

                                        foreach (DataGridViewRow dr in stagedDataGridView.Rows)
                                        {
                                            codeValueArray.Add(dr.Cells[0].Value);
                                        }
                                    }
                                }

                                importFormatProgressBar.Value = 40;

                                foreach (var value in clientName)
                                {
                                    tw.WriteLine("Client: " + value);
                                }

                                String reqItem;
                                String dateItem;
                                tw.WriteLine("");

                                int a = 0;

                                tw.WriteLine("");
                                tw.WriteLine("****************************************************");
                                tw.WriteLine("******************ERROR LIST START******************");
                                tw.WriteLine("****************************************************");
                                tw.WriteLine("");

                                tw.WriteLine("--Required Field Check--");
                                importFormatProgressBar.Value = 50;

                                //String reqItem;
                                foreach (Object selecteditem in reqListBox.SelectedItems)
                                {
                                    reqItem = selecteditem as String;
                                    int reqCurIndex = reqListBox.Items.IndexOf(reqItem);
                                    if (reqCurIndex >= 0)
                                    {
                                        tw.WriteLine("Required Column: " + reqItem);

                                        for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                                        {
                                            try
                                            {
                                                var value = importedfileDataGridView.Rows[i].Cells[reqCurIndex].Value.ToString();
                                                if (string.IsNullOrWhiteSpace(value))
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "." + " This column is required and you have a missing value.");
                                                }
                                            }
                                            catch (Exception)
                                            {
                                                // If we have reached this far, then none of the cells were empty.
                                                tw.WriteLine("No NULL values found in column " + "'" + reqItem + "'");
                                            }
                                        }
                                    }
                                }
                                tw.WriteLine("");

                                tw.WriteLine("--Code Check--");
                                importFormatProgressBar.Value = 60;
                                a = 0;
                                foreach (var s in iffidArray)
                                {
                                    a++;

                                    if (fieldsThatAreCodesArrayColumnCount.Contains(a) == true)
                                    {
                                        tw.WriteLine("\nCOLUMN " + a + ": " + s);//this is the header line in the output file
                                        for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)//this is the loop that spits out the errors
                                        {
                                            var value = importedfileDataGridView.Rows[i].Cells[a - 1].Value.ToString();
                                            if (codeValueArray.Contains(value) == false && value != "")
                                            {
                                                tw.WriteLine("Error at line " + (i + 1) + "." + " The value: '" + value + "' from your imported file does not exist in the database.");
                                            }
                                        }
                                    }
                                }
                                tw.WriteLine("");

                                tw.WriteLine("--Max Length Check--");
                                importFormatProgressBar.Value = 70;
                                a = 0;
                                foreach (var s in seqArray)//cycle through every column
                                {
                                    if (maxLengthFieldColumnNumberArray.Contains(s) == true)//if one of the columns has a max length, enter this IF
                                    {
                                        int index = Array.IndexOf(seqArray, s);
                                        for (int j = 0; j < importedfileDataGridView.Columns.Count; j++)
                                        {
                                            if (index == j)
                                            {
                                                a++;
                                                for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)//this is the loop that spits out the errors
                                                {

                                                    var value = importedfileDataGridView.Rows[i].Cells[j].Value.ToString();
                                                    int valueLength = value.Length;
                                                    int maxValueLength = intMaxLengthFieldArrayValue[a - 1];
                                                    if (valueLength > maxValueLength)
                                                    {
                                                        tw.WriteLine("Column: " + s);
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The value: '" + value + "' from your imported file is " + valueLength + " characters long. This is too long.");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                tw.WriteLine("");
                                tw.WriteLine("--Date Format Check--");
                                importFormatProgressBar.Value = 80;

                                foreach (Object selecteditem in dateListBox.SelectedItems)
                                {
                                    dateItem = selecteditem as String;
                                    int dateCurIndex = dateListBox.Items.IndexOf(dateItem);
                                    if (dateComboBox1.Text == "" && dateComboBox2.Text == "" && dateComboBox3.Text == "")
                                    {
                                        MessageBox.Show("Your date format is NULL. Please create a date format using the dropdown menus.");
                                        return;
                                    }
                                    string dateFormat2 = dateFormat.Text.Remove(0, 13);

                                    int dateFormatLength = dateFormat2.Length;
                                    //MessageBox.Show("dateFormat2=" + dateFormat2+ "dateFormatLength="+ dateFormatLength);
                                    if (dateCurIndex >= 0)
                                    {
                                        if (dateFormatLength == 0)
                                        {
                                            MessageBox.Show("Your date format cannot be empty if you are specifying a date column", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                            return;
                                        }

                                        tw.WriteLine("Date Column: " + dateItem);
                                        for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                                        {
                                            var value = importedfileDataGridView.Rows[i].Cells[dateCurIndex].Value.ToString();

                                            if ((importFormatFindNullCheckbox.Checked) & (value == "" || value == null || value == " "))
                                            {
                                                tw.WriteLine("NULL at line " + (i + 1) + ".");
                                            }

                                            if (dateFormat2 == "yyyymmdd" & (value != "" & value != null & value != " "))
                                            {
                                                try
                                                {
                                                    int year = int.Parse(value.Substring(0, 4));
                                                    int month = int.Parse(value.Substring(4, 2));
                                                    int day = int.Parse(value.Substring(6, 2));


                                                    if (year > 2200)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (month > 12)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (month < 01)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (day > 31)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (day < 01)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }
                                                }
                                                catch
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "." + " Unable to parse the date. Make sure that the date is in the format: " + dateFormat2 + ".");

                                                }
                                            }

                                            if (dateFormat2 == "yyyyddmm" & value != "" & value != null & value != " ")
                                            {
                                                try
                                                {
                                                    int year = int.Parse(value.Substring(0, 4));
                                                    int month = int.Parse(value.Substring(6, 2));
                                                    int day = int.Parse(value.Substring(4, 2));

                                                    if (year > 2200)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (month > 12)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (month < 01)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (day > 31)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (day < 01)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }
                                                }
                                                catch
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "." + " Unable to parse the date. Make sure that the date is in the format: " + dateFormat2 + ".");
                                                }
                                            }

                                            if (dateFormat2 == "yyddmm" & value != "" & value != null & value != " ")
                                            {
                                                try
                                                {
                                                    int year = int.Parse(value.Substring(0, 2));
                                                    int month = int.Parse(value.Substring(4, 2));
                                                    int day = int.Parse(value.Substring(2, 2));

                                                    if (year > 22)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (month > 12)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (month < 01)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (day > 31)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (day < 01)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }
                                                }
                                                catch
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "." + " Unable to parse the date. Make sure that the date is in the format: " + dateFormat2 + ".");
                                                }
                                            }

                                            if (dateFormat2 == "yymmdd" & value != "" & value != null & value != " ")
                                            {
                                                try
                                                {
                                                    int year = int.Parse(value.Substring(0, 2));
                                                    int month = int.Parse(value.Substring(2, 2));
                                                    int day = int.Parse(value.Substring(4, 2));

                                                    if (year > 22)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (month > 12)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (month < 01)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (day > 31)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (day < 01)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }
                                                }
                                                catch
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "." + " Unable to parse the date. Make sure that the date is in the format: " + dateFormat2 + ".");
                                                }
                                            }

                                            if (dateFormat2 == "mmddyyyy" & value != "" & value != null & value != " ")
                                            {
                                                try
                                                {
                                                    int year = int.Parse(value.Substring(4, 4));
                                                    int month = int.Parse(value.Substring(0, 2));
                                                    int day = int.Parse(value.Substring(2, 2));

                                                    if (year > 2200)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (month > 12)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (month < 01)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (day > 31)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (day < 01)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }
                                                }
                                                catch
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "." + " Unable to parse the date. Make sure that the date is in the format: " + dateFormat2 + ".");
                                                }
                                            }

                                            if (dateFormat2 == "mmyyyydd" & value != "" & value != null & value != " ")
                                            {
                                                try
                                                {
                                                    int year = int.Parse(value.Substring(2, 4));
                                                    int month = int.Parse(value.Substring(0, 2));
                                                    int day = int.Parse(value.Substring(6, 2));

                                                    if (year > 2200)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (month > 12)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (month < 01)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (day > 31)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }

                                                    if (day < 01)
                                                    {
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                    }
                                                }
                                                catch
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "." + " Unable to parse the date. Make sure that the date is in the format: " + dateFormat2 + ".");
                                                }
                                            }
                                        }
                                    }
                                }
                                tw.WriteLine("");
                                importFormatRowCountToolStripStatusLabel.Text = importformatDataGridView.Rows.Count.ToString();
                                sqlCounterToolStripStatusLabel.Text = stagedDataGridView.Rows.Count.ToString();
                                conn.Close();
                            }
                            catch { return; }
                            conn.Close();
                        }
                        tw.WriteLine("EOF.");
                    }
                }
                importFormatProgressBar.Value = 90;
                importFormatProgressBar.Value = 100;
                MessageBox.Show("Import Format error file has been created. \nLocation: " + path, "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + @">>>   Import Format error file has been created. Location: C:\Program Files (x86)\DataAnalysisTool\Import Format Error Files");
                progressBar1.MarqueeAnimationSpeed = 0;
                Process.Start(path);
            }
        }

        private void fileSweepGoPictureBox_Click(object sender, EventArgs e)
        {
            fileSweepProgressBar.Value = 0;
            fileSweepGoPictureBox.Enabled = false;
            fileSweepRichTextBox.Clear();
            fileSweepRichTextBox.AppendText(Environment.NewLine +
            @"###########################################################################################" + Environment.NewLine +
            @"########################DataAnalysisTool - FileSweep Progress###############################################" + Environment.NewLine +
            @"###########################################################################################" + Environment.NewLine +
            @"Current Date: " + DateTime.Now + System.Environment.NewLine +
            @"Server: " + serverSelect7.Text + System.Environment.NewLine +
            @"Database: " + fileSweepDatabaseComboBox.Text + System.Environment.NewLine +
            @"" + System.Environment.NewLine +
            @"" + System.Environment.NewLine +
            @"NOTE: This will only bring back in file patterns with an FTP server!" + System.Environment.NewLine +
            @"*****************************************************************" + System.Environment.NewLine +
            @"********************RUN PROGRESS********************" + System.Environment.NewLine +
            @"*****************************************************************" + System.Environment.NewLine
            );
            for (int i = 0; i < fileSweepDataGridView.RowCount; i++)
            {
                fileSweepRichTextBox.AppendText((i+1)+". FILE SWEEP: "+ fileSweepDataGridView.Rows[i].Cells[1].Value+"         FILE: "+fileSweepDataGridView.Rows[i].Cells[0].Value + Environment.NewLine);

            }
            //var localFilePath = @"C:\Users\I868538\Desktop\test6um.xlsx";
            //var ftpUsername = "robwar31";
            //var ftpPassword = "pass";
            //using (WebClient client = new WebClient())
            //{
            //    client.Credentials = new NetworkCredential(ftpUsername, ftpPassword);
            //    var path = Path.Combine("ftp.steelcitysites.net/", "favicon.png");
            //    client.UploadFile("ftp://ftp.steelcitysites.net/test6um.xlsx", WebRequestMethods.Ftp.UploadFile, localFilePath);
            //}
            fileSweepGoPictureBox.Enabled = true;
            fileSweepProgressBar.Value = 100;
        }
    }
}