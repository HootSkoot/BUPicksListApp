using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BUPicksList
{
    public partial class Form1 : Form
    {
        private string pathfilename;
        private Excel.Application xlApp;
        private Excel.Workbook xlWorkbook;
        private Excel.Workbook newXLWorkbook;
        private string formula = "=IFERROR(VLOOKUP(RC[-8],'https://mydrive.amat.com/personal/abisheik_mani_contractor_amat_com/Documents/BU Delivery Process/[BU_Pick_Schedule.xlsx]B21-MissingList'!C2:C5,4,FALSE),IFERROR(VLOOKUP(RC[-8],'https://mydrive.amat.com/personal/abisheik_mani_contractor_amat_com/Documents/BU Delivery Process/[BU_Pick_Schedule.xlsx]B72-MissingList'!C2:C5,4,FALSE),VLOOKUP(RC[-8],'https://mydrive.amat.com/personal/abisheik_mani_contractor_amat_com/Documents/BU Delivery Process/[BU_Pick_Schedule.xlsx]B81-MissingList'!C2:C5,4,FALSE)))";
        private int lastUsedRow;
        public Form1()
        {
            InitializeComponent();
            fillRoomBox();
            fillBUBox();



        }

        private void fillRoomBox()
        {
            RoomList.DataSource = Properties.Settings.Default.RoomList.Cast<string>().ToArray();
            
        }
         private void fillBUBox()
        {
            BUListBox.DataSource = Properties.Settings.Default.BUList.Cast<string>().ToArray();
        }
        

        private void CreateButton_Click(object sender, EventArgs e)
        {
            
        }

        private void AddRoomButton_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.RoomList.Add(RoomTextBox.Text);
            Properties.Settings.Default.Save();
            fillRoomBox();
        }

        private void AddBUButton_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.BUList.Add(BUTextBox.Text);
            Properties.Settings.Default.Save();
            fillBUBox();
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            //dlg.ShowDialog();
            
            if (dlg.ShowDialog() == DialogResult.OK) 
            {
                xlApp = new Excel.Application();
                //get name of file and path
                pathfilename = dlg.FileName;
                FileLabel.Text = Path.GetFileName(dlg.FileName);

                //get raw data as a sheet
                object misValue = System.Reflection.Missing.Value;
                Excel.Worksheet xlWorkSheet;
                Excel.Worksheet newxlWorkSheet;
                
                
                xlWorkbook = xlApp.Workbooks.Open(pathfilename);
                //xlWorkSheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
                xlWorkSheet = xlWorkbook.Sheets[1];
                xlWorkSheet.UsedRange.Copy(Missing.Value);
                

                //create base workbook
                newXLWorkbook = xlApp.Workbooks.Add(Missing.Value);
                newxlWorkSheet = newXLWorkbook.Sheets[1];
                newxlWorkSheet.UsedRange.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, misValue, misValue);
                newxlWorkSheet.Cells[1,9] = "Missing?";
                newxlWorkSheet.Cells[5,9] = formula;

                lastUsedRow = newxlWorkSheet.Cells.Find("*", misValue,
                               misValue, misValue,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, misValue, misValue).Row;

                newxlWorkSheet.Range["i2", "i" + lastUsedRow].FormulaR1C1 = formula;
                //newxlWorkSheet.Range["J3", newxlWorkSheet.UsedRange.RowHeight - 1].FormulaR1C1 = formula;
                newxlWorkSheet.Calculate();
                newXLWorkbook.SaveAs("test-file.xlsx");
                xlWorkbook.Close(true,misValue,misValue);
                newXLWorkbook.Close(true, misValue, misValue);

                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkbook);
                Marshal.ReleaseComObject(xlApp);

                


            }
            
        }

        private void createNewFile(string filename)
        {

        }
    }
}
