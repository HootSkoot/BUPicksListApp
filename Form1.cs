using System;
using System.Data;
using System.Collections.Generic;
using OLEDB = System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;

namespace BUPicksList
{
    public partial class Form1 : Form
    {
        private string pathfilename;
        private static string defaultMissingDirectory = "";
        private string formula = "";
        private string formulaModified = "=IFERROR(VLOOKUP(RC[-8],'" + defaultMissingDirectory + "/[BU_Pick_Schedule.xlsx]B21-MissingList'!C2:C5,4,FALSE),IFERROR(VLOOKUP(RC[-8],'" + defaultMissingDirectory + "/[BU_Pick_Schedule.xlsx]B72-MissingList'!C2:C5,4,FALSE),VLOOKUP(RC[-8],'" + defaultMissingDirectory + "/[BU_Pick_Schedule.xlsx]B81-MissingList'!C2:C5,4,FALSE)))";
        private string formulaCopiedSheet = "=IFERROR(VLOOKUP(RC[-8],MissingList!B:E,4,FALSE),IFERROR(VLOOKUP(RC[-8],MissingList!B:E,4,FALSE),VLOOKUP(RC[-8],MissingList!B:E,4,FALSE)))";
        private string masterData = "";
        private string missingFileName = "BU_Pick_Schedule.xlsx";
        private string path;
        private string masterDataPath;
        private string userPathToMissingFile = "";
        private string fullMissingPath = "";
        private string missingFileOneDriveDirectory = "Abisheik Mani --TR-CNTR - BU Delivery Process";
        private string namePath;
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
                //get input file path
                pathfilename = dlg.FileName;
                FileLabel.Text = Path.GetFullPath(dlg.FileName);
                path = Path.GetDirectoryName(dlg.FileName);


                //Create output file name and path
                masterData = "ExpenseDeliveryManagement " + DateTime.Now.ToString("mm-dd-yyyy") + " - " + DateTime.Now.ToString("hh tt") + ".xlsx";
                masterDataPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + masterData;
               
                //find the missing list workbook name, Get the full OneDrive path
                userPathToMissingFile = Environment.GetEnvironmentVariable("HOMEDRIVE") + Environment.GetEnvironmentVariable("HOMEPATH") + Path.DirectorySeparatorChar + "Applied Materials";
                fullMissingPath = Path.Combine(userPathToMissingFile, missingFileOneDriveDirectory);
                var tempBookName = "";
                if (defaultMissingDirectory != "")
                {
                    tempBookName = defaultMissingDirectory;
                }
                else
                {
                    tempBookName = fullMissingPath + Path.DirectorySeparatorChar + missingFileName;
                }
                
                try
                {
                    //Use EPPlus to load both sheets into master file, copy formula and save
                    using (ExcelPackage packTemp = new ExcelPackage(new FileInfo("./template.xlsx")))
                    using (ExcelPackage packList = new ExcelPackage(new FileInfo(pathfilename)))
                    using (ExcelPackage packMissing = new ExcelPackage(new FileInfo(tempBookName)))
                    using (ExcelPackage pckd = new ExcelPackage())
                    {
                        var sheet1 = packTemp.Workbook.Worksheets["MasterData"];
                        var sheet2 = packTemp.Workbook.Worksheets["MissingList"];
                        var sheetList = packList.Workbook.Worksheets[1];
                        var sheetMissing = packMissing.Workbook.Worksheets["B21-MissingList"];
                        sheet1.Cells[1, 1, sheetList.Dimension.End.Row, sheetList.Dimension.End.Column].Value = sheetList.Cells[1, 1, sheetList.Dimension.End.Row, sheetList.Dimension.End.Column].Value;
                        sheet2.Cells[1, 1, sheetMissing.Dimension.End.Row, sheetMissing.Dimension.End.Column].Value = sheetMissing.Cells[1, 1, sheetMissing.Dimension.End.Row, sheetMissing.Dimension.End.Column].Value;
                        sheet1.Cells[1, 9].Value = "Status";
                        var rangeXML = sheet1.Cells[2, sheet1.Dimension.End.Column, sheet1.Dimension.End.Row, sheet1.Dimension.End.Column];

                        pckd.Workbook.CalcMode = ExcelCalcMode.Automatic;

                        pckd.Workbook.Calculate();
                        packTemp.SaveAs(new FileInfo(masterDataPath));
                    }

                    //Briefly open Excel to load results of formulas and save
                    //No other library could store results without opening Excel
                    var app = new Excel.Application();
                    app.Visible = false;
                    Excel.Workbook wb = app.Workbooks.Open(masterDataPath);
                    wb.Save();
                    wb.Close();
                    app.Quit();

                    CreateDailyFolder();
                    
                    
                }
                catch (Exception ex)
                {
                    string templateResults = "Template file found: " + File.Exists("./template.xlsx");
                    string masterDataResults = "Downloaded Master Data location: " + pathfilename;
                    string missingListResults = "Missing List Location: " + tempBookName;
                    MessageBox.Show(templateResults + "\n" + masterDataResults + "\n" + missingListResults + "\n" + ex.Message, "Error Warning", MessageBoxButtons.OK);
                    throw;
                }
                


            }
            
        }

        //uses created masterlist as a database to query, builds useable file
        private void CreateButton_Click(object sender, EventArgs e)
        {
            if (masterDataPath != "")
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
                try
                {
                    using (var package = new ExcelPackage(new FileInfo(masterDataPath)))  //newXLWorkbook = new ClosedXML.Excel.XLWorkbook(new FileStream(masterDataPath,FileMode.OpenOrCreate,FileAccess.ReadWrite))
                    {
                        SheetCollection largeBook = new SheetCollection("Large and Crate", namePath);
                        foreach (var sheet in package.Workbook.Worksheets)
                        {
                            if (sheet.Name == "MissingList" || sheet.Cells[2, 1].Value == null)
                            {
                                sheet.Hidden = eWorkSheetHidden.Hidden;
                            }
                            else
                            {
                                switch (sheet.Name)
                                {
                                    case string name when sheet.Name.Contains("Data"):
                                        sheet.TabColor = Color.White;
                                        break;
                                    case string name when sheet.Name.Contains("Large") && sheet.Name.Contains("Attempted"):
                                        sheet.TabColor = Color.Orange;
                                        largeBook.addSheet(sheet);
                                        break;
                                    case string name when sheet.Name.Contains("Large"):
                                        sheet.TabColor = Color.Green; //ClosedXML.Excel.XLColor.Green
                                        largeBook.addSheet(sheet);
                                        break;
                                    case string name when sheet.Name.Contains("Attempted"):
                                        sheet.TabColor = Color.Yellow;
                                        SheetCollection book = new SheetCollection(sheet.Name,namePath);
                                        book.addSheet(sheet);
                                        book.createWorkbook();
                                        break;
                                    default:
                                        sheet.TabColor = Color.Blue;
                                        break;
                                }
                            }
                        }
                        largeBook.createWorkbook();
                        package.Save();
                    }

                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message + " " + masterDataPath);
                    throw;
                } 
                
                    
                

            }
            CopyMasterFile();
        }

        


        //Create a reuseable OLEDB connection
        private OLEDB.OleDbConnection returnConnection()
        {
            return new OLEDB.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + masterData + ";Extended Properties=\"Excel 12.0 Xml; HDR = YES\"");
        }

        
        /*
         * These set of methods send SQL querys to the excel data, creating reports pages
         * for the attributes described in the apps listboxes (settings)
         */
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
                        cmd.CommandText = @"Create Table " + roomName.Replace(' ','_').Replace('-','_') + "(RfidTagId varchar, Location varchar,BU varchar, BUStaging varchar, RequestedDate varchar,RequestedModifyDate varchar,LocType varchar,PackageType varchar);";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = @"Insert Into [" + roomName.Replace(' ', '_').Replace('-', '_') + "$] Select RfidTagId,Location,BU,BUStaging,RequestedDate,RequestedModifyDate,LocType,PackageType From [MasterData$] Where Status<>'Missing' and (Location NOT LIKE '%RVN%' and Location NOT LIKE '%RCV-Stag%' and Location NOT LIKE '%Attempt%' and Location NOT LIKE '%pallet%') and BUStaging='" + roomName + "' and PackageType NOT IN " + SizeListString() + " and (LocType<>'CUSTOMER STAGING' and LocType<>'delivered') and Location NOT LIKE '%attempt%' Order By BU;";
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
                        cmd.CommandText = @"Create Table " + roomName.Replace(' ', '_').Replace('-', '_') + "_" + BU + "(RfidTagId varchar, Location varchar,BU varchar, BUStaging varchar, RequestedDate varchar,RequestedModifyDate varchar,LocType varchar,PackageType varchar);";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = @"Insert Into [" + roomName.Replace(' ', '_').Replace('-', '_') + "_" + BU + "$] Select RfidTagId,Location,BU,BUStaging,RequestedDate,RequestedModifyDate,LocType,PackageType From [MasterData$] Where Status<>'Missing' and (Location NOT LIKE '%RVN%' and Location NOT LIKE '%RCV-Stag%' and Location NOT LIKE '%Attempt%' and Location NOT LIKE '%pallet%') and BUStaging='" + roomName + "' and BU='" + BU + "' and PackageType NOT IN " + SizeListString() + " and (LocType<>'CUSTOMER STAGING' and LocType<>'delivered') and Location NOT LIKE '%attempt%' Order By BU;";
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
                        cmd.CommandText = @"Create Table " + roomName.Replace(' ', '_').Replace('-', '_') + "_Large(RfidTagId varchar, Location varchar,BU varchar, BUStaging varchar, RequestedDate varchar,RequestedModifyDate varchar,LocType varchar,PackageType varchar);";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = @"Insert Into [" + roomName.Replace(' ', '_').Replace('-', '_') + "_Large$] Select RfidTagId,Location,BU,BUStaging,RequestedDate,RequestedModifyDate,LocType,PackageType From [MasterData$] Where Status<>'Missing' and (Location NOT LIKE '%RVN%' and Location NOT LIKE '%RCV-Stag%' and Location NOT LIKE '%Attempt%' and Location NOT LIKE '%B72%' and BUStaging NOT LIKE '%B72%') and BUStaging='" + roomName + "' and PackageType IN " + SizeListString() + " and (LocType<>'CUSTOMER STAGING' and LocType<>'delivered') Order By BU;";
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
                        cmd.CommandText = @"Create Table " + buildingName.Replace(' ', '_').Replace('-', '_') + "_Attempted(RfidTagId varchar, Location varchar,BU varchar, BUStaging varchar, RequestedDate varchar,RequestedModifyDate varchar,LocType varchar,PackageType varchar);";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = @"Insert Into [" + buildingName.Replace(' ', '_').Replace('-', '_') + "_Attempted$] Select RfidTagId,Location,BU,BUStaging,RequestedDate,RequestedModifyDate,LocType,PackageType From [MasterData$] Where Status<>'Missing' and PackageType NOT IN " + SizeListString() + " and (LocType<>'CUSTOMER STAGING' and LocType<>'delivered') and (Location NOT LIKE '%versum%' and Location NOT LIKE '%MarkGruver%' and (Location LIKE '%" + buildingName + "%attempt%' OR (Location LIKE '%" + buildingName + "%pallet%' AND LocType LIKE '%ATTEMPT%') OR (BUStaging LIKE '%" + buildingName + "%' AND LocType Like '%ATTEMPT%'))) Order By BU;";
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
                        cmd.CommandText = @"Create Table " + roomName.Replace(' ', '_').Replace('-', '_') + "_Large_Attempted(RfidTagId varchar, Location varchar,BU varchar, BUStaging varchar, RequestedDate varchar,RequestedModifyDate varchar,LocType varchar,PackageType varchar);";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = @"Insert Into [" + roomName.Replace(' ', '_').Replace('-', '_') + "_Large_Attempted$] Select RfidTagId,Location,BU,BUStaging,RequestedDate,RequestedModifyDate,LocType,PackageType From [MasterData$] Where Status<>'Missing' and (LocType<>'CUSTOMER STAGING' and LocType<>'delivered') and PackageType IN " + SizeListString() + " and (Location NOT LIKE '%RVN%' and Location NOT LIKE '%RCV-Stag%' and Location NOT LIKE '%versum%' and (Location LIKE '%attempt%' OR LocType LIKE '%ATTEMPT%') and Location NOT LIKE '%B72%' and BUStaging NOT LIKE '%B72%') and BUStaging='" + roomName + "' Order By BU;";
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

        //replaces the live missing list if network unavailable
        public void UserSelectedMissingList()
        {
            OpenFileDialog dlg = new OpenFileDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                defaultMissingDirectory = dlg.FileName;
                MissingListLabel.Text = Path.GetDirectoryName(dlg.FileName);
                formula = formulaModified;
            }
        }
        private void CreateDailyFolder()
        {
            namePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + "BUPicks " + DateTime.Today.ToString("MM-dd-yy");
            if (!Directory.Exists(namePath))
            {
                Directory.CreateDirectory(namePath);
            }
        }

        private void CopyMasterFile()
        {
            string destination = Path.Combine(namePath, masterData);
            File.Copy(masterDataPath, destination, true);
        }
    }
}
