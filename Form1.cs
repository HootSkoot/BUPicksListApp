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
using ClosedExcel = ClosedXML.Excel;
using System.Net;

namespace BUPicksList
{
    public partial class Form1 : Form
    {
        private string pathfilename;
        private Excel.Application xlApp;
        private ClosedXML.Excel.XLWorkbook xlWorkbook;
        private ClosedXML.Excel.XLWorkbook newXLWorkbook;
        private static string defaultMissingDirectory = "https://mydrive.amat.com/personal/abisheik_mani_contractor_amat_com/Documents/BU Delivery Process/";
        private string formula = "=IFERROR(VLOOKUP(RC[-8],'https://mydrive.amat.com/personal/abisheik_mani_contractor_amat_com/Documents/BU Delivery Process/[BU_Pick_Schedule.xlsx]B21-MissingList'!C2:C5,4,FALSE),IFERROR(VLOOKUP(RC[-8],'https://mydrive.amat.com/personal/abisheik_mani_contractor_amat_com/Documents/BU Delivery Process/[BU_Pick_Schedule.xlsx]B72-MissingList'!C2:C5,4,FALSE),VLOOKUP(RC[-8],'https://mydrive.amat.com/personal/abisheik_mani_contractor_amat_com/Documents/BU Delivery Process/[BU_Pick_Schedule.xlsx]B81-MissingList'!C2:C5,4,FALSE)))";
        private string formulaModified = "=IFERROR(VLOOKUP(RC[-8],'" + defaultMissingDirectory + "/[BU_Pick_Schedule.xlsx]B21-MissingList'!C2:C5,4,FALSE),IFERROR(VLOOKUP(RC[-8],'" + defaultMissingDirectory + "/[BU_Pick_Schedule.xlsx]B72-MissingList'!C2:C5,4,FALSE),VLOOKUP(RC[-8],'" + defaultMissingDirectory + "/[BU_Pick_Schedule.xlsx]B81-MissingList'!C2:C5,4,FALSE)))";
        private int lastUsedRow;
        private string masterData = "";
        private string path;
        private string masterDataPath;
        private string missingFilePath = "https://mydrive.amat.com/:x:/r/personal/abisheik_mani_contractor_amat_com/Documents/BU%20Delivery%20Process/BU_Pick_Schedule.xlsx?d=wc1badf9411494be8bc10fe4c675e6b14&csf=1&web=1&e=YlpIFx";
        public Form1()
        {
            InitializeComponent();
            fillRoomBox();
            fillBUBox();
            fillBuildingBox();
            fillSizeBox();


        }
        //Use settings to save desired inputs, Datasource for each list is the setting
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

        private void fillSizeBox()
        {
            SizeListBox.DataSource = Properties.Settings.Default.SizeList.Cast<string>().ToArray();
        }
        
        //uses created masterlist as a database to query, builds useable file
        private void CreateButton_Click(object sender, EventArgs e)
        {
            if (masterData != "")
            {
                foreach (String roomName in Properties.Settings.Default.RoomList)
                {
                    Console.WriteLine(roomName);
                    if (roomName.Contains("Pong"))
                    {
                        foreach (String BU in Properties.Settings.Default.BUList)
                        {
                            CreateNormalSheet(roomName, BU);
                        }
                    }
                    else
                    {
                        CreateNormalSheet(roomName);
                    }
                    
                }
                foreach (String buildingName in Properties.Settings.Default.BuildingList)
                {
                    Console.WriteLine(buildingName);
                    CreateAttemptedSheet(buildingName);
                }
                foreach (String roomName in Properties.Settings.Default.RoomList)
                {
                    if (roomName.Contains("Dock"))
                    {
                        Console.WriteLine(roomName);
                        CreateLargeSheet(roomName);
                    }
                }
                foreach (String roomName in Properties.Settings.Default.RoomList)
                {
                    if (roomName.Contains("Dock"))
                    {
                        Console.WriteLine(roomName);
                        CreateLargeAttemptedSheet(roomName);
                    }
                }
                using (var workbook = new ClosedExcel.XLWorkbook(masterDataPath))
                {
                    foreach (var sheet in workbook.Worksheets)
                    {
                        if (sheet.LastRowUsed().RowNumber() == 1)
                        {
                            sheet.Hide();
                        }
                        else
                        {
                            switch (sheet.Name)
                            {
                                case string name when sheet.Name.Contains("Data"):
                                    sheet.TabColor = ClosedExcel.XLColor.White;
                                    break;
                                case string name when sheet.Name.Contains("Large"):
                                    sheet.TabColor = ClosedExcel.XLColor.Green;
                                    break;
                                case string name when sheet.Name.Contains("Attempted"):
                                    sheet.TabColor = ClosedExcel.XLColor.Yellow;
                                    break;
                                case string name when sheet.Name.Contains("Large") && sheet.Name.Contains("Attempted"):
                                    sheet.TabColor = ClosedExcel.XLColor.Orange;
                                    break;
                                default:
                                    sheet.TabColor = ClosedExcel.XLColor.Blue;
                                    break;
                            }
                        }
                    }
                    workbook.Save();
                }
            }
            
           
            
        }
        //adds the user entered string to the desired setting
        private void AddRoomButton_Click(object sender, EventArgs e)
        {
            if (RoomTextBox.Text.Trim() != "")
            {
                Properties.Settings.Default.RoomList.Add(RoomTextBox.Text);
                Properties.Settings.Default.Save();
                fillRoomBox();
            }
        }

        private void AddBUButton_Click(object sender, EventArgs e)
        {
            if (BUTextBox.Text.Trim() != "")
            {
                Properties.Settings.Default.BUList.Add(BUTextBox.Text);
                Properties.Settings.Default.Save();
                fillBUBox();
            }
        }

        private void AddBuildingButton_Click(object sender, EventArgs e)
        {
            if (BuildingTextBox.Text.Trim() != "")
            {
                Properties.Settings.Default.BuildingList.Add(BuildingTextBox.Text);
                Properties.Settings.Default.Save();
                fillBuildingBox();
            }
            
        }

        private void SizeButton_Click(object sender, EventArgs e)
        {
            if (SizeTextBox.Text.Trim() != "")
            {
                Properties.Settings.Default.SizeList.Add(SizeTextBox.Text);
                Properties.Settings.Default.Save();
                fillSizeBox();
            }
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();

            if (dlg.ShowDialog() == DialogResult.OK) 
            {
                
                pathfilename = dlg.FileName;
                FileLabel.Text = Path.GetDirectoryName(dlg.FileName);
                path = Path.GetDirectoryName(dlg.FileName);

                //get raw data as a sheet
                object misValue = System.Reflection.Missing.Value;
                xlWorkbook = new ClosedXML.Excel.XLWorkbook(pathfilename);
                var xlWorkSheet = xlWorkbook.Worksheet(1);
                var firstCell = xlWorkSheet.FirstCellUsed();
                var lastCell = xlWorkSheet.LastCellUsed();
                var range = xlWorkSheet.Range(firstCell.Address, lastCell.Address);



                //create base workbook
                masterData = "ExpenseDeliveryManagement " + DateTime.Now.ToString("mm-dd-yyyy") + " - " + DateTime.Now.ToString("hh tt") + ".xlsx";
                var newXLWorkbook = new ClosedXML.Excel.XLWorkbook();
                var newxlWorkSheet = newXLWorkbook.Worksheets.Add("MasterData");
                newxlWorkSheet.Cell(1, 1).Value = range;
                MissingList();
                //var missingList = new ClosedExcel.XLWorkbook("https://mydrive.amat.com/personal/abisheik_mani_contractor_amat_com/Documents/BU Delivery Process/BU_Pick_Schedule.xlsx");
                //missingList.Worksheet(1).CopyTo(newXLWorkbook, "B21-MissingList");


                //insert desired formula

                newxlWorkSheet.Cell(1,9).Value = "Status";
                firstCell = newxlWorkSheet.FirstRowUsed().RowBelow().LastCellUsed().CellRight();
                lastCell = newxlWorkSheet.LastRowUsed().LastCellUsed().CellRight();
                range = newxlWorkSheet.Range(firstCell.Address, lastCell.Address);
                range.FormulaR1C1 = formula;

                //Ensure that the sheet formulas have resolved
                masterDataPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + masterData;
                newXLWorkbook.CalculateMode = ClosedExcel.XLCalculateMode.Auto;
                newXLWorkbook.SaveAs(masterDataPath);
                foreach (var cell in range.CellsUsed())
                {
                    var result = newxlWorkSheet.Cell(cell.Address).CachedValue;
                    cell.Value = result;
                }




                //Sheet needs to be saved and closed to reset objects, turn the formulas into flat values

                //ReplaceFormulasWithValues();
                
                
            }
            
        }

        private void MissingList()
        {
            FileStream fileStream = null;
            FileStream fstream = null;
            try
            {
                //missingFilePath = missingFilePath.Replace( "/", "\\");
                //missingFilePath = missingFilePath.Replace("https:", "");
                //missingFilePath = missingFilePath.Replace(" ", "%20");
                //fileStream = new MemoryStream(File.ReadAllBytes(missingFilePath));
                String fileName = Path.GetFileName(missingFilePath);
                //Stream stream = File.Open(missingFilePath, FileMode.Open);
                byte[] data;
                byte[] buffer = new byte[2048];
                WebRequest request = WebRequest.Create(missingFilePath);
                using (WebResponse response = request.GetResponse())
                {
                    using (Stream responseStream = response.GetResponseStream())
                    {
                        using (MemoryStream ms = new MemoryStream())
                        {
                            int count = 0;
                            do
                            {
                                count = responseStream.Read(buffer, 0, buffer.Length);
                                ms.Write(buffer, 0, count);
                            } while (count != 0);
                            data = ms.ToArray();
                        }
                    }
                }
                string filePath = "Docs";
                using (fstream = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite))
                {
                    fstream.Write(data, 0, data.Length);
                    fstream.Close();
                }
                
            }
            catch (FileFormatException ex)
            {
                //null;
            }
            
            
        }
        

        //Create a reuseable OLEDB connection
        private OLEDB.OleDbConnection returnConnection()
        {
            return new OLEDB.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + masterData + ";Extended Properties=\"Excel 12.0 Xml; HDR = YES\"");
        }

        private void ReplaceFormulasWithValues()
        {
            /*
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
            */
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
                        cmd.CommandText = @"Insert Into [" + roomName.Replace(' ', '_').Replace('-', '_') + "$] Select * From [MasterData$] Where Status<>'Missing' and (Location NOT LIKE '%RVN%' and Location NOT LIKE '%RCV-Stag%' and Location NOT LIKE '%Attempt%') and BUStaging='" + roomName + "' and PackageType NOT IN " + SizeListString() + " and (LocType<>'CUSTOMER STAGING' and LocType<>'delivered') and Location NOT LIKE '%attempt%' Order By BU;";
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

        private void CreateNormalSheet(String roomName, String BU)
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
                        cmd.CommandText = @"Create Table " + roomName.Replace(' ', '_').Replace('-', '_') + "_" + BU + "(RfidTagId varchar, Location varchar,BU varchar, BUStaging varchar, RequestedDate varchar,RequestedModifyDate varchar,LocType varchar,PackageType varchar, Status varchar);";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = @"Insert Into [" + roomName.Replace(' ', '_').Replace('-', '_') + "_" + BU + "$] Select * From [MasterData$] Where Status<>'Missing' and (Location NOT LIKE '%RVN%' and Location NOT LIKE '%RCV-Stag%' and Location NOT LIKE '%Attempt%') and BUStaging='" + roomName + "' and BU='" + BU + "' and PackageType NOT IN " + SizeListString() + " and (LocType<>'CUSTOMER STAGING' and LocType<>'delivered') and Location NOT LIKE '%attempt%' Order By BU;";
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

        private void CreateLargeSheet(String roomName)
        {
            try
            {

                using (OLEDB.OleDbConnection conn = returnConnection())
                {
                    try
                    {
                        conn.Open();
                        OLEDB.OleDbCommand cmd = new OLEDB.OleDbCommand();
                        cmd.Connection = conn;
                        cmd.CommandText = @"Create Table " + roomName.Replace(' ', '_').Replace('-', '_') + "_Large(RfidTagId varchar, Location varchar,BU varchar, BUStaging varchar, RequestedDate varchar,RequestedModifyDate varchar,LocType varchar,PackageType varchar, Status varchar);";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = @"Insert Into [" + roomName.Replace(' ', '_').Replace('-', '_') + "_Large$] Select * From [MasterData$] Where Status<>'Missing' and (Location NOT LIKE '%RVN%' and Location NOT LIKE '%RCV-Stag%' and Location NOT LIKE '%Attempt%') and BUStaging='" + roomName + "' and PackageType IN " + SizeListString() + " and (LocType<>'CUSTOMER STAGING' and LocType<>'delivered') Order By BU;";
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


        private void CreateAttemptedSheet(String buildingName)
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
                        cmd.CommandText = @"Create Table " + buildingName.Replace(' ', '_').Replace('-', '_') + "_Attempted(RfidTagId varchar, Location varchar,BU varchar, BUStaging varchar, RequestedDate varchar,RequestedModifyDate varchar,LocType varchar,PackageType varchar, Status varchar);";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = @"Insert Into [" + buildingName.Replace(' ', '_').Replace('-', '_') + "_Attempted$] Select * From [MasterData$] Where Status<>'Missing' and PackageType NOT IN " + SizeListString() + " and (LocType<>'CUSTOMER STAGING' and LocType<>'delivered') and (Location NOT LIKE '%versum%' and Location NOT LIKE '%MarkGruver%' and Location LIKE '%" + buildingName + "%attempt%') Order By BU;";
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

        private void CreateLargeAttemptedSheet(String roomName)
        {
            try
            {

                using (OLEDB.OleDbConnection conn = returnConnection())
                {
                    try
                    {
                        conn.Open();
                        OLEDB.OleDbCommand cmd = new OLEDB.OleDbCommand();
                        cmd.Connection = conn;
                        cmd.CommandText = @"Create Table " + roomName.Replace(' ', '_').Replace('-', '_') + "_Large_Attempted(RfidTagId varchar, Location varchar,BU varchar, BUStaging varchar, RequestedDate varchar,RequestedModifyDate varchar,LocType varchar,PackageType varchar, Status varchar);";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = @"Insert Into [" + roomName.Replace(' ', '_').Replace('-', '_') + "_Large_Attempted$] Select * From [MasterData$] Where Status<>'Missing' and (LocType<>'CUSTOMER STAGING' and LocType<>'delivered') and PackageType IN " + SizeListString() + " and (Location NOT LIKE '%RVN%' and Location NOT LIKE '%RCV-Stag%' and Location NOT LIKE '%versum%' and Location LIKE '%attempt%') and BUStaging='" + roomName + "' Order By BU;";
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

        //Returns a string form of the large size list
        private string SizeListString()
        {
            string value = "(";
            int count = 0;

            if (Properties.Settings.Default.SizeList.Count != 0)
            {
                foreach (String size in Properties.Settings.Default.SizeList)
                {
                    value += "'" + size + "'";
                    count += 1;
                    if (count != Properties.Settings.Default.SizeList.Count)
                    {
                        value += ",";
                    }
                }
                value += ")";
            }
            else
            {
                value = "('Large (> 35 lbs)','Crate')";
            }

            return value;
        }

        private void DeleteRoomButton_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.RoomList.Remove(RoomList.SelectedItem.ToString());
            Properties.Settings.Default.Save();
            fillRoomBox();
        }

        private void RemoveBUButton_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.BUList.Remove(BUListBox.SelectedItem.ToString());
            Properties.Settings.Default.Save();
            fillBUBox();
        }

        private void RemoveBuildingButton_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.BuildingList.Remove(BuildingListBox.SelectedItem.ToString());
            Properties.Settings.Default.Save();
            fillBuildingBox();
        }

        private void RemoveSizeButton_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.SizeList.Remove(SizeListBox.SelectedItem.ToString());
            Properties.Settings.Default.Save();
            fillSizeBox();
        }

        public void LocateMissingButton_Click(object sender, EventArgs e)
        {
            UserSelectedMissingList();
        }

        public void UserSelectedMissingList()
        {
            OpenFileDialog dlg = new OpenFileDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                defaultMissingDirectory = Path.GetDirectoryName(dlg.FileName);
                MissingListLabel.Text = Path.GetDirectoryName(dlg.FileName);
                formula = formulaModified;
            }
        }
    }
}
