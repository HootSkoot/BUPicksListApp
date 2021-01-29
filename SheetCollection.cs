using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace BUPicksList
{
    /// <summary>
    /// Generic Sheet Collection and Workbook Creation
    /// stores sheets and creates a workbook out of sheets at given location
    /// </summary>
    class SheetCollection
    {
        List<ExcelWorksheet> sheets;
        string title;
        string namePath;

        public SheetCollection(string _namePath)
        {
            sheets = new List<ExcelWorksheet>();
            setName("");
            setPath(_namePath);
        }

        public SheetCollection(string title, string _namePath)
        {
            sheets = new List<ExcelWorksheet>();
            setName(title);
            setPath(_namePath);
        }

        public void setName(string title)
        {
            this.title = title;
        }

        public void setPath(string _namePath)
        {
            namePath = _namePath;
        }

        public void addSheet(ExcelWorksheet sheet)
        {
            sheets.Add(sheet);
        }
        /// <summary>
        /// Creates a Workbook at the given directory
        /// </summary>
        public void createWorkbook()
        {
            if (sheets.Count > 0)
            {
                using (ExcelPackage pck = new ExcelPackage())
                {
                    foreach (var sheet in sheets)
                    {
                        pck.Workbook.Worksheets.Add(sheet.Name, sheet);

                    }
                    pck.SaveAs(new System.IO.FileInfo(namePath + "\\" + title + " " + DateTime.Now.ToString("HH-tt") + ".xlsx"));
                }
            }
            
        }

    }
}
