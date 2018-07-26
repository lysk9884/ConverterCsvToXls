using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ConverterCsvToXls.scripts
{
    public class CsvToXls
    {
        private static CsvToXls _instance = null;
        public char Delimiter = ',';
        public string EndofLine = "\r\n";

        public static CsvToXls GetInstance
        {
            get
            {
                if (_instance == null) _instance = new CsvToXls();
                return _instance;
            }
        }

        public void Convert(string csvPath, string xlsPath)
        {
            var fileContent = File.ReadAllLines(@csvPath);

            var modifiedContent = string.Empty;
            string tempQuote = string.Empty;

            var maxCol = 0;

            List<int> dateColumnNumbers = new List<int>();
            Dictionary<int, Dictionary<int, string>> dataTables = new Dictionary<int, Dictionary<int, string>>();
            var dataColumn = 0;
            var dataRow = 0;
            List<string> cells = new List<string>();

            for (int lineIdx = 0; lineIdx < fileContent.Length; lineIdx++)
            {
                var line = fileContent[lineIdx];
                var lineWords = line.Split(',');

                for (int wordIdx = 0; wordIdx < lineWords.Length; wordIdx++)
                {
                    var cell = lineWords[wordIdx];

                    if ((!cell.Contains('\"') && string.IsNullOrEmpty(tempQuote)) || string.IsNullOrEmpty(cell))
                    {
                        cells.Add(cell);
                    }
                    else // Quotation 이  있을때
                    {
                        tempQuote += cell;

                        if (cell.ElementAt(cell.Length - 1) == '\"') // Quotation 이 끝났을때 Cells 에 추가 한다.
                        {
                            tempQuote = tempQuote.Replace("\"", "");
                            cells.Add(tempQuote);
                            tempQuote = string.Empty;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(tempQuote))
                {
                    tempQuote += "\n";
                }

                if (lineIdx == 0) maxCol = cells.Count;
            }

            for (int i = 0; i < cells.Count; i++)
            {
                if(!dataTables.ContainsKey(dataRow)) dataTables.Add(dataRow , new Dictionary<int , string>());
                dataTables[dataRow].Add(dataColumn, cells[i]);

                modifiedContent += cells[i];

                if ((i + 1) % maxCol == 0) // 한 행이 끝났을때 End Of Line 을 넣어준다.
                {
                    dataRow += 1;
                    dataColumn = 0;
                    modifiedContent += EndofLine;
                }
                else
                {
                    dataColumn += 1;
                    modifiedContent += Delimiter;
                }
            }

            if (File.Exists(xlsPath)) File.Delete(xlsPath);
            ExcelPackage package = new ExcelPackage(new FileInfo(xlsPath));
            string worksheetsName = Path.GetFileNameWithoutExtension(csvPath);
            var format = new ExcelTextFormat();
            format.Delimiter = Delimiter;
            format.EOL = EndofLine;
            format.Encoding = Encoding.UTF8;

            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetsName);
            worksheet.Cells["A1"].LoadFromText(modifiedContent, format);
            worksheet.Cells[worksheet.Dimension.Start.Row, worksheet.Dimension.Start.Column, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].Style.Numberformat.Format = "@";

            for (int column = worksheet.Dimension.Start.Column; column <= worksheet.Dimension.End.Column; column++)
            {
                for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
                {
                    if (worksheet.Cells[row, column].Value != null)
                    {
                        worksheet.Cells[row, column].Style.Numberformat.Format = "@";
                        worksheet.Cells[row, column].Value = dataTables[row - 1][column - 1];
                    }
                }
            }

            File.Delete(csvPath);
            package.Save();
            package.Dispose();
        }
    }
}