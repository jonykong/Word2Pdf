using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace Word2Pdf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // 选择文件
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;//该值确定是否可以选择多个文件
            dialog.Title = "请选择doc";
            dialog.Filter = "PPT |*.docx;*doc";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string file = dialog.FileName;
                this.textBox1.ResetText();
                this.textBox1.AppendText(file);
            }
        }

        // 开始
        private void button2_Click(object sender, EventArgs e)
        {
            string filePath = this.textBox1.Text;
           
            if (filePath.EndsWith(".doc") || filePath.EndsWith(".docx"))
            {
                Process(filePath);
                string filedir = Path.GetDirectoryName(filePath);
                System.Diagnostics.Process.Start("explorer.exe", filedir);
            }
        }
        public static void Process(string file)
        {
            Console.WriteLine("Found: {0}", file);
            File.SetAttributes(file, FileAttributes.Normal);
            var idx = file.LastIndexOf(".");
            var newPath = file.Remove(idx, file.Length - idx) + ".pdf";
            Doc2PDFAtServerClass.word2PdfFcih(file, newPath);
        }
    }
}
