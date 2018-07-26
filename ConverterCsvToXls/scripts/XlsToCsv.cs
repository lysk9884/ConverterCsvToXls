using OfficeOpenXml;
using System.IO;
using System.Text;

namespace ConverterCsvToXls.scripts
{
    public class XlsToCsv
    {
        public char Delimiter = ',';
        public string EndofLine = "\r\n";

        private static XlsToCsv _instance = null;

        public static XlsToCsv GetInstance
        {
            get
            {
                if (_instance == null) _instance = new XlsToCsv();
                return _instance;
            }
        }

        public void Convert(string xlsPath, string csvPath)
        {
            ExcelPackage package = new ExcelPackage(new FileInfo(xlsPath));
            string worksheetsName = Path.GetFileNameWithoutExtension(csvPath);
            var format = new ExcelTextFormat();
            format.Delimiter = Delimiter;
            format.EOL = EndofLine;
            format.Encoding = Encoding.UTF8;

            string csvText = string.Empty;

            foreach (var worksheet in package.Workbook.Worksheets)
            {
                for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
                {
                    for (int column = worksheet.Dimension.Start.Column; column <= worksheet.Dimension.End.Column; column++)
                    {
                        if (worksheet.Cells[row, column].Value != null)
                        {
                            var strValue = worksheet.Cells[row, column].Value.ToString();

                            if (strValue.Contains(",") || strValue.Contains("\n")) // 가 있을때 Quotation 이라고 생각해서 "" 를 삽입한다.
                            {
                                csvText += string.Format("\"{0}\"", strValue);
                            }
                            else
                            {
                                csvText += strValue;
                            }

                            if(column != worksheet.Dimension.End.Column) csvText += Delimiter; // 마지막 열에는 붙일 필요가 없다.
                        }
                    }
                    csvText += EndofLine;
                }
            }

            package.Dispose();
            File.Delete(xlsPath);

            if (File.Exists(csvPath)) File.Delete(csvPath);
            var sw = new StreamWriter(csvPath, false, Encoding.UTF8);
            sw.WriteLine(csvText);
            sw.Close();
        }
    }
}