using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace EJTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Open file
            string[] fileArray = Directory.GetFiles(tbEJPath.Text);
            string sBankName = BankName.GetItemText(BankName.SelectedItem);
            if(sBankName == "BAB-Opteva")
            {
                ProcessOptevaEJ processEJ = new ProcessOptevaEJ();
                processEJ.ProcessOptevaEJFile(fileArray);
            }
            else if (sBankName == "SHB-CS280")
            {
                ProcessProcashATMEJ processProcashATMEJ = new ProcessProcashATMEJ();
                processProcashATMEJ.ProcessProcashATMEJFiles(fileArray);
            }
            else
            {
                ProcessEJ processEJ = new ProcessEJ();
                processEJ.ProcessEJFiles(fileArray);
            }
            
            MessageBox.Show("Done!");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string folderPath = "";
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                folderPath = folderBrowserDialog1.SelectedPath;
                tbEJPath.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void BankName_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
