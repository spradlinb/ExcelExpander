using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelOpp = Microsoft.Office.Interop.Excel;

namespace ExcelExpander
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("Enter full path of Excel file: ");
            var inboundPath = Console.ReadLine();

            Console.Write("Enter folder location for final file: ");
            var outboundPath = Console.ReadLine();

            ExcelOpp.Application xlApp = new ExcelOpp.Application();

            ExcelOpp.Workbook xlWorkbook = xlApp.Workbooks.Open(inboundPath);
            ExcelOpp._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            ExcelOpp.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            //rowCount = 30;

            Dictionary<int, int> colCounts = new Dictionary<int, int>();
            List<string> headerRow = new List<string>();

            for (int i = 1; i <= rowCount; i++)
            {
                var colValues = new List<string>();
                for (int j = 1; j <= colCount; j++)
                {
                    if (xlRange.Cells[i, j].Value2 != null)
                    {
                        var cellValue = xlRange.Cells[i, j].Value2.ToString();
                        colValues.Add(cellValue);
                        if (i == 1)
                            colCounts.Add(j, 1); // Add initial values to colCounts dictionary
                        else
                        {
                            if (!string.IsNullOrEmpty(cellValue))
                            {
                                var cellSplit = cellValue.Split(',');
                                var cellSplitCount = cellSplit.Length;
                                if (cellSplitCount > colCounts[j])
                                    colCounts[j] = cellSplitCount;
                            }
                        }
                    }
                }
                if (i == 1)
                    headerRow = colValues; // Set the header row values
                Console.WriteLine("Checking counts in row " + i);
            }

            var allData = new List<string>();
            //allData.Add(string.Join(",", headerRow.ToArray()));

            var expandedHeader = "";
            for (int x = 0; x <= colCount-1; x++)
            {
                if (colCounts[x+1] > 1)
                {
                    for (int y = 1; y <= colCounts[x+1]; y++)
                    {
                        expandedHeader += headerRow[x] + " " + y + ",";
                    }
                }
                else
                {
                    expandedHeader += headerRow[x] + ",";
                }
            }
            allData.Add(expandedHeader);

            for (int i = 2; i <= rowCount; i++)
            {
                var rowData = "";
                for (int j = 1; j <= colCount; j++)
                {
                    var cellValue = "";
                    if (xlRange.Cells[i, j].Value2 != null)
                    {
                        cellValue = xlRange.Cells[i, j].Value2.ToString();
                    }

                    var expandedColCount = colCounts[j];
                    var splitCellValuesCount = cellValue.Split(',').Count();
                    var extraCommas = expandedColCount - splitCellValuesCount;

                    rowData += cellValue + ",";
                    for (int x = 1; x <= extraCommas; x++)
                        rowData += ",";
                }
                Console.WriteLine("Creating expanded row " + i);
                allData.Add(rowData);
            }

            outboundPath += "\\ExpandedFile.csv";
            File.WriteAllLines(outboundPath, allData.ToArray());
            Console.Write("File conversion complete. Press any key to finish.");
            Console.ReadLine();

        }
    }
}
