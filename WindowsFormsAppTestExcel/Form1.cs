using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsAppTestExcel
{
    public partial class Form1 : Form
    {
        //object[,] arrData; //объявляем двумерный массив (строки и столбцы)

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //имя Excel файла  
            string fileNameAddin = textBoxAddinPath.Text;
            string fileNameExcelWorkbook = textBoxExcelSheetPath.Text;

            Cursor savedCursor = this.Cursor;

            // Application
            Excel.Application xlApp = new Excel.Application();

            try
            {
                this.Cursor = Cursors.WaitCursor;

                // открываем Excel файлы
                Excel.Workbook xlWb1 = xlApp.Workbooks.Open(fileNameAddin);
                Excel.Workbook xlWb2 = xlApp.Workbooks.Open(fileNameExcelWorkbook);

                Excel.Worksheet xlSht = xlWb2.Sheets["Лист1"]; //имя листа в файле
                xlApp.Visible = true;
                //xlApp.DisplayAlerts = false;
                xlApp.Run(@"ParserAddinTest");
                string activeParserName = textBoxActiveParserName.Text;
                var res = xlApp.Run(@"StartParser", activeParserName);
                if (!String.IsNullOrEmpty(res))
                {
                    throw new Exception(res);
                }

                int iLastRow = xlSht.Cells[xlSht.Rows.Count, "B"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А
                var arrData = (object[,])xlSht.Range["A1:B" + iLastRow].Value; //берём данные со 2-й строки, если нужно с 1-й, то замените A2 на A1

                NAR(xlSht);
                xlWb1.Close(false);
                NAR(xlWb1);
                xlWb2.Close(false); //закрываем файл и сохраняем изменения, если не сохранять, то false                
                NAR(xlWb2);

                StringBuilder sb = new StringBuilder();

                for (int i = 1; i <= arrData.GetUpperBound(0); i++) //заполняем ComboBox данными из массива
                {
                    //this.comboBox1.Items.Add(arrData[i, 1]);
                    sb.Append(arrData[i, 2]);
                    sb.Append("\r\n");
                }
                textBoxResult.Text = sb.ToString();
            }
            catch (Exception exc)
            {
                textBoxResult.Text = exc.Message;
            }
            finally
            {
                
                xlApp.Quit(); //закрываем Excel
                NAR(xlApp);
                GC.Collect();
                this.Cursor = savedCursor;
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string path = LoadParserPath();
            textBoxAddinPath.Text = path;

            textBoxActiveParserName.Text = LoadParserDefaultSetup();
        }

        private string LoadParserPath()
        {
            const string regRoot = @"HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Parser\Setup";
            string s = (string)Registry.GetValue(regRoot, "AddinPath", "") ?? "";

            return s;
        }

        private string LoadParserDefaultSetup()
        {
            const string regRoot = @"HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Parser\Settings";
            string s = (string)Registry.GetValue(regRoot, "ACTIVE_PARSER", "") ?? "";

            return s;

        }

        private void btnParserTest_Click(object sender, EventArgs e)
        {
            string fileNameParser = textBoxAddinPath.Text;
            Cursor savedCursor = this.Cursor;
            try
            {
                this.Cursor = Cursors.WaitCursor;
                textBoxResult.Text = "Parser test starting. Please wait...";
                // Application
                Excel.Application xlApp = new Excel.Application();

                // открываем Excel файл
                Excel.Workbook xlWb1 = xlApp.Workbooks.Open(fileNameParser);

                xlApp.Run(@"ParserAddinTest");
                xlWb1.Close(false);
                NAR(xlWb1);
                xlApp.Quit(); //закрываем Excel
                NAR(xlApp);
                GC.Collect();
                textBoxResult.Text = "Parser test ok. xlApp = null";
            }
            catch (Exception exc)
            {
                textBoxResult.Text = exc.Message;
            }
            finally
            {
                this.Cursor = savedCursor;
            }
        }

        private void btnSelectExcelFile_Click(object sender, EventArgs e)
        {
            if (openExcelFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxExcelSheetPath.Text = openExcelFileDialog.FileName;
            }
        }

        private static void NAR(object o)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(o);
            }
            catch { }
            finally
            {
                o = null;
            }
        }
    }
}
