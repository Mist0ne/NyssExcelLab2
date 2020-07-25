using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Windows;
using System.Collections.ObjectModel;

namespace NyssExcelLab2
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel()
        {
            // Base constructor. Usually used while trying create a new file
        }

        public Excel(string path, int Sheet)
        {
            // Constructor. Input data: path to the Excel file, number of sheet in this file
            this.path = path;

            wb = excel.Workbooks.Open(path); // Create WorkBook
            ws = wb.Worksheets[Sheet]; // Create WorkSheet
        }

        ~Excel()
        {
            excel.Workbooks.Close();
        }

        public string ReadCell(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].Value2 != null)
                return ws.Cells[i, j].Value2.ToString();
            else
                return "";
        } // Read one cell from the Excel file

        public string[,] ReadRange(int starti, int starty, int endi, int endy)
        {
            starti++;
            starty++;
            endi++;
            endy++;

            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            object[,] holder = range.Value2;
            string[,] returnstring = new string[endi - starti, endy - starty];
            for (int i = 1; i <= endi-starti; i++)
            {
                for (int j = 1; j <= endy - starty; j++)
                {
                    try
                    {
                        returnstring[i - 1, j - 1] = holder[i, j].ToString();
                    }
                    catch
                    {
                        return returnstring;
                    }
                }
            }
            return returnstring;
        } // Read range of cells from the Excel file

        public void WriteToCell(int i, int j, string s)
        {
            i++;
            j++;
            ws.Cells[i, j].Value2 = s;
        } // Write string in the cell

        public void WriteRange(int starti, int starty, int endi, int endy, ObservableCollection<ThreatModel> data)
        {
            string[,] convertedData = new string[data.Count, 8];
            for (int i = 0; i < data.Count; i++)
            {
                convertedData[i, 0] = data[i].Id;
                convertedData[i, 1] = data[i].Name;
                convertedData[i, 2] = data[i].Capture;
                convertedData[i, 3] = data[i].ThreatSource;
                convertedData[i, 4] = data[i].ThreatTarget;
                convertedData[i, 5] = data[i].Confidentiality;
                convertedData[i, 6] = data[i].Integrity;
                convertedData[i, 7] = data[i].Availability;
            }

            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            range.Value2 = convertedData;
        } // Write data from ObservableCollection in range of cells in the Excel file

        public void SaveAs(string _path)
        {
            ws.Activate();
            ws.SaveAs(Filename: _path);
        } // Save active WorkBook with Sheets in Excel file (which transmitted as _path)

        public void NewFile()
        {
            this.wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.ws = wb.Worksheets[1];
        } // Creating a new file (WorkBook and WorkSheet)

        public static string GetPath()
        {
            string[] path_to_exe = Directory.GetCurrentDirectory().Split('\\');
            string right_path = "";
            for (int i = 0; i < path_to_exe.Length - 2; i++)
            {
                right_path += path_to_exe[i] + '\\';
            }
            return right_path;
        } // Return string with path to the program directory (using '\' as separator)
    }
}
