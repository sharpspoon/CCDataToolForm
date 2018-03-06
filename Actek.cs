using System;
using System.Globalization;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Diagnostics;
using System.Text;
using System.Xml;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using System.DirectoryServices.ActiveDirectory;
using System.Web;

namespace CCDataTool
{
    public partial class acteksoft : Form
    {
        public acteksoft()
        {
            InitializeComponent();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            var processdir = Environment.CurrentDirectory; 
            System.IO.Directory.CreateDirectory(processdir + @"\ACTEKSOFT");
            string path = processdir + @"\ACTEKSOFT\LaunchWithACTEKSOFT.cmd";
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
            {
                using (TextWriter tw = new StreamWriter(fs))
                {
                    //tw.WriteLine(@"md C:\Program Files(x86)\CCDataTool\Data");
                    //tw.WriteLine("robocopy "+processdir+" "+@"""C:\Program Files (x86)\CCDataTool\Data"""+@" /MIR");
                    tw.WriteLine("taskkill /IM CCDataImportTool.exe /F");
                    tw.WriteLine("cls");
                    tw.WriteLine(@"C:\Windows\System32\runas.exe /user:ACTEKSOFT\"+textBox7.Text+ @" /netonly " +@""""+processdir+@"\CCDataImportTool.exe""");
                    tw.WriteLine("exit");
                }
                System.Diagnostics.Process.Start(path);
            }
        }
        private void acteksoft_Load(object sender, EventArgs e)
        {
        }
    }
}
