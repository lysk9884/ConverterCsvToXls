using ConverterCsvToXls.scripts;
using System;
using System.IO;

namespace ConverterCsvToXls
{
    public partial class Converter : System.Web.UI.Page
    {
        public string _csvResultDir = @"C:\Temp\Csv\";
        public string _xlsResultDir = @"C:\Temp\Xls\";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(XlsDirInput.Value)) XlsDirInput.Value = _xlsResultDir;
            if (string.IsNullOrEmpty(CsvDirInput.Value)) CsvDirInput.Value = _csvResultDir;
        }

        protected void ToXlsBtnClicked(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty( XlsDirInput.Value)) _xlsResultDir = XlsDirInput.Value;
            if (!Directory.Exists(_xlsResultDir)) Directory.CreateDirectory(_xlsResultDir);

            for (int i = 0; i < Request.Files.Count; i++)
            {
                var file = Request.Files[i];

                if (file.ContentLength > 0)
                {
                    var csvFilePath = Path.GetFileName(file.FileName);
                    csvFilePath = _xlsResultDir + csvFilePath;
                    file.SaveAs(csvFilePath);
                    var fileName = Path.GetFileNameWithoutExtension(csvFilePath);
                    var fileExt = Path.GetExtension(csvFilePath);
                    var xlsFilePath = csvFilePath.Replace(fileExt, ".xlsx");

                    CsvToXls.GetInstance.Convert(csvFilePath, xlsFilePath);
                }
            }
        }

        protected void ToCsvBtnClicked(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(CsvDirInput.Value)) _csvResultDir = CsvDirInput.Value;
            if (!Directory.Exists(_xlsResultDir)) Directory.CreateDirectory(_csvResultDir);

            for (int i = 0; i < Request.Files.Count; i++)
            {
                var file = Request.Files[i];

                if (file.ContentLength > 0)
                {
                    var xlsFilePath = Path.GetFileName(file.FileName);
                    xlsFilePath = _csvResultDir + xlsFilePath;
                    file.SaveAs(xlsFilePath);
                    var fileName = Path.GetFileNameWithoutExtension(xlsFilePath);
                    var fileExt = Path.GetExtension(xlsFilePath);
                    var csvFilePah = xlsFilePath.Replace(fileExt, ".csv");
                    XlsToCsv.GetInstance.Convert(xlsFilePath, csvFilePah);
                }
            }
        }
    }
}