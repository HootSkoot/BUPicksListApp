using System;
using System.Data;
using OLEDB = System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
//using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
//using System.Data.OleDb;

namespace BUPicksList
{
    public partial class Form1 : Form
    {
        private string pathfilename;
        private Excel.Application xlApp;
        private Excel.Workbook xlWorkbook;
        private Excel.Workbook newXLWorkbook;
        private Excel.Worksheet xlWorkSheet;
        private Excel.Worksheet newxlWorkSheet;
        private string formula = "=IFERROR(VLOOKUP(RC[-8],'https://mydrive.amat.com/personal/abisheik_mani_contractor_amat_com/Documents/BU Delivery Process/[BU_Pick_Schedule.xlsx]B21-MissingList'!C2:C5,4,FALSE),IFERROR(VLOOKUP(RC[-8],'https://mydrive.amat.com/personal/abisheik_mani_contractor_amat_com/Documents/BU Delivery Process/[BU_Pick_Schedule.xlsx]B72-MissingList'!C2:C5,4,FALSE),VLOOKUP(RC[-8],'https://mydrive.amat.com/personal/abisheik_mani_contractor_amat_com/Documents/BU Delivery Process/[BU_Pick_Schedule.xlsx]B81-MissingList'!C2:C5,4,FALSE)))";
        private int lastUsedRow;
        private string masterData = "";
        private string path;
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

        private void fillBuildingBox()
        {
            BuildingListBox.DataSource = Properties.Settings.Default.BuildingList.Cast<string>().ToArray();
        }
        

        private void CreateButton_Click(object sender, EventArgs e)
        {
            if (masterData != "")
            {
                foreach (String roomName in Properties.Settings.Default.RoomList)
                {
                    Console.WriteLine(roomName);
                    CreateNormalSheet(roomName);
                }
            }
            
           
            //createNewFile("testing");
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
                path = Path.GetDirectoryName(dlg.FileName);

                //get raw data as a sheet
                object misValue = System.Reflection.Missing.Value;
                //Excel.Worksheet xlWorkSheet;
                //Excel.Worksheet newxlWorkSheet;
                
                
                xlWorkbook = xlApp.Workbooks.Open(pathfilename);
                //xlWorkSheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
                xlWorkSheet = xlWorkbook.Sheets[1];
                xlWorkSheet.UsedRange.Copy(Missing.Value);
                

                //create base workbook
                newXLWorkbook = xlApp.Workbooks.Add(Missing.Value);
                newxlWorkSheet = newXLWorkbook.Sheets[1];
                newxlWorkSheet.Name = "MasterData";
                newxlWorkSheet.UsedRange.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, misValue, misValue);
                newxlWorkSheet.Cells[1,9] = "Status";
                newxlWorkSheet.Cells[5,9] = formula;

                lastUsedRow = newxlWorkSheet.Cells.Find("*", misValue,
                               misValue, misValue,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, misValue, misValue).Row;

                newxlWorkSheet.Range["i2", "i" + lastUsedRow].FormulaR1C1 = formula;
                //newxlWorkSheet.Range["J3", newxlWorkSheet.UsedRange.RowHeight - 1].FormulaR1C1 = formula;

                newxlWorkSheet.Calculate();
                //Testing copying over formulas
                //'J' is the column with the formula inserted
                


                //Excel.Range range = (Excel.Range)newxlWorkSheet.Range["i2"].EntireColumn;
                //newxlWorkSheet.Range["i2", "i" + lastUsedRow].PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues,Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                //range.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, Missing.Value, Missing.Value);
                //end test

                masterData = "ExpenseDeliveryManagement " + DateTime.Now.ToString("mm-dd-yyyy") + " - " + DateTime.Now.ToString("hh tt") + ".xlsx";
                newXLWorkbook.SaveAs(masterData);
                xlWorkbook.Close(true,misValue,misValue);
                newXLWorkbook.Close(true, misValue, misValue);

                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkbook);
                Marshal.ReleaseComObject(newXLWorkbook);
                Marshal.ReleaseComObject(newxlWorkSheet);
                Marshal.ReleaseComObject(xlApp);

                ReplaceFormulasWithValues();




            }
            
        }

        

        private void createSheet(string BUname, Excel.Workbook wkbook)
        {
            //wksheet.EnableAutoFilter = true;
            Excel.Worksheet newsheet = wkbook.Sheets.Add(After: wkbook.Sheets[wkbook.Sheets.Count]);
            Excel.Worksheet oldsheet = wkbook.Sheets[1];
            Excel.Range range = oldsheet.UsedRange;
            //range.AutoFilter(Field: 9, Criteria1: "#N/A", Operator: XlAutoFilterOperator.xlOr, Criteria2: "Located");
            range.Copy(Missing.Value);
            newsheet.UsedRange.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, Missing.Value, Missing.Value);

            //Excel.Range sourceRange = oldsheet.UsedRange;
            //sourceRange.AutoFilter("9", "#N/A", Excel.XlAutoFilterOperator.xlAnd, "Located", false);

            //Excel.Range filteredRange = sourceRange.SpecialCells(XlCellType.xlCellTypeVisible, XlSpecialCellsValue.xlTextValues);

            //filteredRange.Copy(newsheet);
            newsheet.Name = "Test";
            Marshal.ReleaseComObject(newsheet);
            Marshal.ReleaseComObject(oldsheet);
        }

        private OLEDB.OleDbConnection returnConnection()
        {
            return new OLEDB.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + masterData + ";Extended Properties=\"Excel 12.0 Xml; HDR = YES\"");
        }

        private DataSet loadMasterSheet(string fileName, string sheetName)
        {
            DataSet sheetData = new DataSet();
            using (OLEDB.OleDbConnection conn = this.returnConnection())
            {
                conn.Open();
                // retrieve the data using data adapter
                OLEDB.OleDbDataAdapter sheetAdapter = new OLEDB.OleDbDataAdapter("select * from [" + sheetName + "]", conn);
                sheetAdapter.Fill(sheetData);
            }
            return sheetData;
        }

        private void ReplaceFormulasWithValues()
        {
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(masterData);
            xlWorkSheet = xlWorkbook.Sheets[1];
            Excel.Range range = (Excel.Range)xlWorkSheet.UsedRange.Columns["I:I",Type.Missing];
            range.Copy(Type.Missing);
            range.PasteSpecial(Excel.XlPasteType.xlPasteValues,Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            //standard close out of INterop
            xlWorkbook.Save();
            xlWorkbook.Close(true, Missing.Value, Missing.Value);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkbook);
            //Marshal.ReleaseComObject(newXLWorkbook);
            //Marshal.ReleaseComObject(newxlWorkSheet);
            Marshal.ReleaseComObject(xlApp);
        }

        private void CreateNormalSheet(String roomName)
        {
            try
            {
                //Fetches sheet with master list data

                using (OLEDB.OleDbConnection conn = returnConnection())
                {
                    try
                    {
                        conn.Open();
                        OLEDB.OleDbCommand cmd = new OLEDB.OleDbCommand();
                        cmd.Connection = conn;
                        cmd.CommandText = @"Create Table " + roomName.Replace(' ','_').Replace('-','_') + "(RfidTagId varchar, Location varchar,BU varchar, BUStaging varchar, RequestedDate varchar,RequestedModifyDate varchar,LocType varchar,PackageType varchar, Status varchar);";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = @"Insert Into [" + roomName.Replace(' ', '_').Replace('-', '_') + "$] Select * From [MasterData$] Where Status<>'Missing' and BUStaging='" + roomName + "' and (PackageType NOT LIKE '%Large%' and PackageType NOT LIKE '%Crate%') and (LocType<>'CUSTOMER STAGING' and LocType<>'delivered') and Location NOT LIKE '%attempt%' Order By BU;";
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                    }
                    finally
                    {
                        conn.Close();
                        conn.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void CreateAttemptedSheet(String roomName)
        {
            try
            {
                //Fetches sheet with master list data

                using (OLEDB.OleDbConnection conn = returnConnection())
                {
                    try
                    {
                        conn.Open();
                        OLEDB.OleDbCommand cmd = new OLEDB.OleDbCommand();
                        cmd.Connection = conn;
                        cmd.CommandText = @"Create Table " + roomName.Replace(' ', '_').Replace('-', '_') + "Attempted(RfidTagId varchar, Location varchar,BU varchar, BUStaging varchar, RequestedDate varchar,RequestedModifyDate varchar,LocType varchar,PackageType varchar, Status varchar);";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = @"Insert Into [" + roomName.Replace(' ', '_').Replace('-', '_') + "$] Select * From [MasterData$] Where Status<>'Missing' and BUStaging='" + roomName + "' and (PackageType NOT LIKE '%Large%' and PackageType NOT LIKE '%Crate%') and (LocType<>'CUSTOMER STAGING' and LocType<>'delivered') and Location LIKE '%attempt%' Order By BU;";
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                    }
                    finally
                    {
                        conn.Close();
                        conn.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
