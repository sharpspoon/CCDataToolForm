using System;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Linq;

namespace SAPDataAnalysisTool
{
    public partial class SAPDataAnalysisTool
    {
         /*
         * ############################################################################################   
         * ############################################################################################
         * ####################BENCHMARK TAB###########################################################
         * ############################################################################################
         * ############################################################################################
        */
        //------------------GO------------------------------------------------------
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
    }
}