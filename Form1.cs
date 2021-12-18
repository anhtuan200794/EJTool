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
            string[] fileArray = Directory.GetFiles(tbEJPath.Text, "*.jrn");
            ProcessEJ processEJ = new ProcessEJ();
            processEJ.ProcessEJFiles(fileArray);
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
    }
}
