using System;
using System.Data;
using System.Windows.Forms;
using OfficeExcel = Microsoft.Office.Interop.Excel;

namespace ExportToExcel
{
    public partial class frmExcelExport : Form
    {
        /// <summary>
        /// Returns a dataset already populated with test data
        /// </summary>
        /// <returns></returns>
        private DataSet GetDataSet()
        {
            DataSet ds = new DataSet();
            DataTable dtbl = new DataTable();
            dtbl.Columns.Add("Sl No");//Define Columns
            dtbl.Columns.Add("Novel Name");
            dtbl.Columns.Add("Author");
            dtbl.Columns.Add("Genres");
            dtbl.Columns.Add("Published Date");
            dtbl.Columns.Add("Price");
            dtbl.Columns.Add("Rating");

            dtbl.Rows.Add("1", "In Search of Lost Time", "Marcel Proust", "Literary modernism", "01-01-1913", "348", "4.3");//Adding Rows
            dtbl.Rows.Add("2", "Ulysses", "James Joyce", "Modernism", "22-02-1922", "58", "3.7");
            dtbl.Rows.Add("3", "Moby Dick", "Herman Melville", "Adventure fiction", "18-10-1851", "131", "3.4");
            dtbl.Rows.Add("4", "Hamlet", "William Shakespeare", "Tragedy", "01-01-1603", "225", "3.9");
            dtbl.Rows.Add("5", "War and Peace", "Leo Tolstoy", "Historical fiction", "01-01-1869", "133.95", "4.1");
            dtbl.TableName = "Table1";
            ds.Tables.Add(dtbl);

            DataTable dtbl2 = dtbl.Copy();//Created copies of first table
            dtbl2.TableName = "Table2";
            ds.Tables.Add(dtbl2);
            DataTable dtbl3 = dtbl.Copy();//Created copies of first table
            dtbl3.TableName = "Table3";
            ds.Tables.Add(dtbl3);

            return ds;
        }

        /// <summary>
        /// Fuction to export dataset to excel
        /// </summary>
        /// <param name="ds"></param>
        private void ExportDataSetToExcel(DataSet ds, string strPath)
        {
            int inHeaderLength = 3, inColumn = 0, inRow = 0;
            System.Reflection.Missing Default = System.Reflection.Missing.Value;
            //Create Excel File
            strPath += @"\Excel" + DateTime.Now.ToString().Replace(':', '-') + ".xlsx";
            OfficeExcel.Application excelApp = new OfficeExcel.Application();
            OfficeExcel.Workbook excelWorkBook = excelApp.Workbooks.Add(1);
            foreach (DataTable dtbl in ds.Tables)
            {
                //Create Excel WorkSheet
                OfficeExcel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add(Default, excelWorkBook.Sheets[excelWorkBook.Sheets.Count], 1, Default);
                excelWorkSheet.Name = dtbl.TableName;//Name worksheet

                //Write Column Name
                for (int i = 0; i < dtbl.Columns.Count; i++)
                    excelWorkSheet.Cells[inHeaderLength + 1, i + 1] = dtbl.Columns[i].ColumnName.ToUpper();

                //Write Rows
                for (int m = 0; m < dtbl.Rows.Count; m++)
                {
                    for (int n = 0; n < dtbl.Columns.Count; n++)
                    {
                        inColumn = n + 1;
                        inRow = inHeaderLength + 2 + m;
                        excelWorkSheet.Cells[inRow, inColumn] = dtbl.Rows[m].ItemArray[n].ToString();
                        if (m % 2 == 0)
                            excelWorkSheet.get_Range("A" + inRow.ToString(), "G" + inRow.ToString()).Interior.Color = System.Drawing.ColorTranslator.FromHtml("#FCE4D6");
                    }
                }

                //Excel Header
                OfficeExcel.Range cellRang = excelWorkSheet.get_Range("A1", "G3");
                cellRang.Merge(false);
                cellRang.Interior.Color = System.Drawing.Color.White;
                cellRang.Font.Color = System.Drawing.Color.Gray;
                cellRang.HorizontalAlignment = OfficeExcel.XlHAlign.xlHAlignCenter;
                cellRang.VerticalAlignment = OfficeExcel.XlVAlign.xlVAlignCenter;
                cellRang.Font.Size = 26;
                excelWorkSheet.Cells[1, 1] = "Greate Novels Of All Time";

                //Style table column names
                cellRang = excelWorkSheet.get_Range("A4","G4");
                cellRang.Font.Bold = true;
                cellRang.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                cellRang.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#ED7D31");
                excelWorkSheet.get_Range("F4").EntireColumn.HorizontalAlignment = OfficeExcel.XlHAlign.xlHAlignRight;
                //Formate price column
                excelWorkSheet.get_Range("F5").EntireColumn.NumberFormat = "0.00";
                //Auto fit columns
                excelWorkSheet.Columns.AutoFit();
            }

            //Delete First Page
            excelApp.DisplayAlerts = false;
            Microsoft.Office.Interop.Excel.Worksheet lastWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkBook.Worksheets[1];
            lastWorkSheet.Delete();
            excelApp.DisplayAlerts = true;

            //Set Defualt Page
            (excelWorkBook.Sheets[1] as OfficeExcel._Worksheet).Activate();

            excelWorkBook.SaveAs(strPath, Default, Default, Default, false, Default, OfficeExcel.XlSaveAsAccessMode.xlNoChange, Default, Default, Default, Default, Default);
            excelWorkBook.Close();
            excelApp.Quit();

            MessageBox.Show("Excel generated successfully \n As "+strPath);
        }
        public frmExcelExport()
        {
            InitializeComponent();
        }
        private void btnExport_Click(object sender, EventArgs e)
        {
            DataSet dsData = GetDataSet();
            ExportDataSetToExcel(dsData, Application.StartupPath);
        }
    }
}
