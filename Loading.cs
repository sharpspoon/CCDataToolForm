using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataAnalysisTool
{
    public partial class Loading : Form
    {
        public Loading()
        {

            InitializeComponent();
            this.BackColor = Color.LimeGreen;
            this.TransparencyKey = Color.LimeGreen;
            

        }

        private void Loading_Load(object sender, EventArgs e)
        {

            //this.Close();
        }

        private void Loading_Leave(object sender, EventArgs e)
        {
            //this.Close();
        }

        private void Loading_VisibleChanged(object sender, EventArgs e)
        {
            //this.Close();
        }

        private void Loading_Deactivate(object sender, EventArgs e)
        {
            //this.Close();
        }
    }
}
