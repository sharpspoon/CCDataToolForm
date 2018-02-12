using System;
using System.Globalization;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using System.DirectoryServices.ActiveDirectory;
using System.Web;

namespace CCDataImportTool
{
    class Medicare
    {

        internal void button2_Click()
        {
            MessageBox.Show("NULL value found in column #3 (PBP)  at line  This is a required field.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
        }
    }
}
