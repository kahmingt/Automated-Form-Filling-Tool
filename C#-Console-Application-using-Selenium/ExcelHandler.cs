using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace RPAChallange
{
    public class ExcelHandler
    {
        private static Excel.Application _application;
        private static Excel.Workbook _workbook;
        private static Excel.Sheets _sheets;
        private Dictionary<string, string> _cell = new();
        private List<Dictionary<string, string>> _rowDetails = new();
        private List<List<Dictionary<string, string>>> _worksheetRows = new();

        ~ExcelHandler()
        {
            Marshal.ReleaseComObject(_sheets);
            Marshal.ReleaseComObject(_workbook);
            Marshal.ReleaseComObject(_application);
        }



        private int GetLastRowIndex(Excel.Worksheet worksheet)
        {
            return worksheet.Cells.Find(
                What: "*",
                SearchOrder: Excel.XlSearchOrder.xlByRows,
                SearchDirection: Excel.XlSearchDirection.xlPrevious,
                MatchCase: false
            ).Row + 1;
        }

        private int GetLastColIndex(Excel.Worksheet worksheet)
        {
            return worksheet.Cells.Find(
                What: "*",
                SearchOrder: Excel.XlSearchOrder.xlByColumns,
                SearchDirection: Excel.XlSearchDirection.xlPrevious,
                MatchCase: false
            ).Column + 1;
        }


        public List<List<Dictionary<string, string>>> ReadAndExtractData(string path)
        {
            _application = new Excel.Application();
            _application.Visible = false;
            _workbook = _application.Workbooks.Open(path);

            // Assume:
            // 1. WorkSheets have row header.
            // 2. Well formatted.
            // 3. Read Left to Right, Top to Bottom
            foreach (Excel.Worksheet worksheet in _workbook.Worksheets)
            {
                //Console.WriteLine("[I]: Reading worksheet [" + worksheet.Name + "]");
                _rowDetails = new();

                for (int row = 2; row < GetLastRowIndex(worksheet); row++)
                {
                    _cell = new();

                    for (int col = 1; col < GetLastColIndex(worksheet); col++)
                    {
                        var header = worksheet.Cells[1, col].Value.ToString().Replace(" ", "").Trim().ToUpper();
                        var value = worksheet.Cells[row, col].Value.ToString().Trim();
                        //Console.WriteLine("[I]: Row " + row + " _cell[" + header + "]=" + value);
                        _cell[header] = value;
                    }
                    
                    _rowDetails.Add(_cell);
                }
                _worksheetRows.Add(_rowDetails);
            }

            _application.Quit();
            return _worksheetRows;
        }

    }
}
