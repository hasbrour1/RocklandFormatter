using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

/*
 * @author: Robert Hasbrouck
 * @version: 1.3
 */

namespace TestApp1
{
    public partial class Form1 : Form
    {

        public static Excel._Workbook tmpBook = null;
        public static Excel.Application tmpApp = null;
        public static Excel._Worksheet tmpSheet = null;
        public static Excel._Workbook srcBook = null;
        public static Excel.Application srcApp = null;
        public static Excel._Worksheet srcSheet = null;

        private string filename = "";   // Output file name	

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void folderpath_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                String filepath = openFileDialog1.FileName;
                this.label1.Text = "Opening " + filepath;
                Execute(filepath);
                this.label1.Text = "Complete";

            }
        }
        
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
           
        }

        //Execute Method takes the file selected and formats it 
        //to the rockland template specification.  The new excel 
        //workbook is then saved in the same directory as the file
        //that was selected.
        public void Execute(String path)
        {

            filename = path;

            //templatePath will be where template.xls will be stored
            //source path get from browse button
            //appPath is directory the .exe is located, which the template should be in same directory
            string appPath = Application.StartupPath;
            string templatePath, sourcePath;

            //root of exe plus template name
            appPath += "\\template.xls";
            
            templatePath = @appPath;
            sourcePath = @filename;

            //If the template file is found, run program
            if (File.Exists(templatePath) && File.Exists(sourcePath))
            {
                //Template setup
                tmpApp = new Excel.Application();
                tmpApp.Visible = false;
                tmpBook = (Excel._Workbook)(tmpApp.Workbooks.Add(templatePath));
                tmpSheet = (Excel._Worksheet)tmpBook.Sheets[1];

                //source setup
                srcApp = new Excel.Application();
                srcApp.Visible = false;
                srcBook = (Excel._Workbook)(srcApp.Workbooks.Add(sourcePath));
                srcSheet = (Excel._Worksheet)srcBook.Sheets[1];

                //Merge Both files here
                for (int i = 1; i < 33; i++)
                {
                    string str = srcSheet.Cells[2, i].Value2;
                    tmpSheet.Cells[3, i] = str;

                }

                //Results page
                tmpSheet = (Excel._Worksheet)tmpBook.Sheets[2];
                srcSheet = (Excel._Worksheet)srcBook.Sheets[2];

                int mod = 1;

                for (int i = 2; i <= srcSheet.UsedRange.Rows.Count; i++)
                {
                    //skip FLDTMP Row
                    if (srcSheet.Cells[i, 1].Value2 != "FLDTEMP")
                    {
                        for (int j = 1; j < 16; j++)
                        {
                            string str = srcSheet.Cells[i, j].Value2;
                            tmpSheet.Cells[i + mod, j] = str;
                        }
                        //re size row height 
                        tmpSheet.Cells[i + 1, 1].RowHeight = 12.75;
                    }
                    else
                    {
                        mod = 0;
                    }
                }

                object m_objOpt = System.Reflection.Missing.Value;

                //remove extension of filename
                string[] remove = { ".xls", ".xlsx"};

                foreach (string item in remove)
                    if (filename.EndsWith(item))
                    {
                        filename = filename.Substring(0, filename.LastIndexOf(item));
                        break; //only allow one match at most
                    }

                //save workbook
                tmpBook.SaveAs(@filename + "_Formatted.xls", m_objOpt,
                        m_objOpt, m_objOpt, true, m_objOpt, Excel.XlSaveAsAccessMode.xlNoChange,
                        m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt);
                tmpBook.Close(true);
                tmpApp.Workbooks.Close();
                tmpApp.Quit();
                srcBook.Close(true);
                srcApp.Workbooks.Close();
                srcApp.Quit();  
                //Open file location after saved
                if (File.Exists(@filename + "_Formatted.xls"))
                {
                    Process.Start("explorer.exe", "/select, " + @filename + "_Formatted.xls");
                }

            }
            else
            {
                this.label1.Text = "Error: File not found.  Please Try again.";
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
