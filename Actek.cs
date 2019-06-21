using System;
using System.Windows.Forms;
using System.IO;

namespace SAPDataAnalysisTool
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
                    tw.WriteLine("taskkill /IM SAPDataAnalysisTool.exe /F");
                    tw.WriteLine("cls");
                    tw.WriteLine(@"C:\Windows\System32\runas.exe /user:ACTEKSOFT\"+textBox7.Text+ @" /netonly " +@""""+processdir+ @"\SAPDataAnalysisTool.exe""");
                    tw.WriteLine("exit");
                }
                System.Diagnostics.Process.Start(path);
            }
        }
        private void acteksoft_Load(object sender, EventArgs e)
        {
        }

        private void toolStripMenuItemClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
