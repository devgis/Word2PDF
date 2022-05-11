using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace Word2PDF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "word|*.docx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                tbFilePath.Text = ofd.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (File.Exists(tbFilePath.Text))
            {
                //Microsoft.Office.Interop.Word.ApplicationClass appClass=new ApplicationClass();
                ////Microsoft.Office.Interop.Word. app=new Microsoft.Office.Interop.Word.ApplicationClass();
                //Microsoft.Office.Interop.Word.Document doc=appClass.
                
                SaveFileDialog sfd=new SaveFileDialog();
                sfd.Filter="pdf|*.pdf";
                if(sfd.ShowDialog()==DialogResult.OK)
                {
                    object FilePath=tbFilePath.Text;
                    object SavePath=sfd.FileName;
                    object saveFormart=WdSaveFormat.wdFormatPDF;

                    Object Nothing = System.Reflection.Missing.Value;

                    Microsoft.Office.Interop.Word.Application WordApp = new Microsoft.Office.Interop.Word.Application();//创建word应用程序对象
                    Microsoft.Office.Interop.Word.Document WordDoc = WordApp.Documents.Open(ref FilePath, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);


                    WordDoc.SaveAs2(ref SavePath, ref saveFormart, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
                    MessageBox.Show("OK");
                }
               

                
            }
            else
            {
                MessageBox.Show("文件不存在！");
            }
        }
    }
}
