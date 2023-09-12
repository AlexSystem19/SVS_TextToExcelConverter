using System.IO;
using System.Linq;
using SVS_TextToExcelConverter_3;
using Excel = Microsoft.Office.Interop.Excel;

namespace SVS_TextToExcelConverter_3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ConvertTextToExcel();
        }

        //private void ConvertTextToExcel()
        //{
        //    string[] filePaths = Directory.GetFiles(@"D:\LBB-GUIAS\", "*.txt");

        //    Excel.Application xlApp = new Excel.Application();
        //    Excel.Workbook xlWorkbook = xlApp.Workbooks.Add();
        //    Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
        //    Excel.Range xlRange = xlWorksheet.UsedRange;

        //    int row = 1;
        //    foreach (string filePath in filePaths)
        //    {
        //        string[] lines = File.ReadAllLines(filePath);
        //        foreach (string line in lines)
        //        {
        //            string[] data = line.Split(';');
        //            for (int i = 0; i < data.Length; i++)
        //            {
        //                xlWorksheet.Cells[row, i + 1] = data[i];
        //            }
        //            row++;
        //        }
        //    }

        //    xlWorkbook.SaveAs(@"D:\LBB-GUIAS\output.xlsx");
        //    xlWorkbook.Close();
        //    xlApp.Quit();
        //}

        //private void ConvertTextToExcel()
        //{
        //    string[] filePaths = Directory.GetFiles(@"D:\LBB-GUIAS\", "*.txt");

        //    Excel.Application xlApp = new Excel.Application();
        //    Excel.Workbook xlWorkbook = xlApp.Workbooks.Add();
        //    Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];

        //    // Set column names for the Excel worksheet
        //    xlWorksheet.Cells[1, 1] = "GUIA";
        //    xlWorksheet.Cells[1, 2] = "FECHA DE EMISION";
        //    xlWorksheet.Cells[1, 3] = "FECHA DE TRASLADO";
        //    xlWorksheet.Cells[1, 4] = "TRACTO";
        //    xlWorksheet.Cells[1, 5] = "REMOLQUE";
        //    xlWorksheet.Cells[1, 6] = "CONDUCTOR";
        //    xlWorksheet.Cells[1, 7] = "LICENCIA";
        //    xlWorksheet.Cells[1, 8] = "CARGA";
        //    xlWorksheet.Cells[1, 9] = "GLOSA";
        //    xlWorksheet.Cells[1, 10] = "EMPRESA";
        //    xlWorksheet.Cells[1, 11] = "CONTENEDOR";
        //    xlWorksheet.Cells[1, 12] = "TICKET DE PESAJE";
        //    xlWorksheet.Cells[1, 13] = "PESO BRUTO";
        //    xlWorksheet.Cells[1, 14] = "PESO TARA";
        //    xlWorksheet.Cells[1, 15] = "PESO NETO";

        //    int row = 2;
        //    foreach (string filePath in filePaths)
        //    {
        //        string[] lines = File.ReadAllLines(filePath);
        //        foreach (string line in lines)
        //        {
        //            string[] data = line.Split(';');

        //            // Write each value to the appropriate cell in the Excel worksheet
        //            if (data.Length >= 12)
        //            {
        //                xlWorksheet.Cells[row, 1] = data[0];
        //                xlWorksheet.Cells[row, 2] = data[1];
        //                xlWorksheet.Cells[row, 3] = data[2];
        //                xlWorksheet.Cells[row, 4] = data[3];
        //                xlWorksheet.Cells[row, 5] = data[4];
        //                xlWorksheet.Cells[row, 6] = data[5] + " " + data[6];
        //                xlWorksheet.Cells[row, 7] = data[7];
        //                xlWorksheet.Cells[row, 8] = data[8];
        //                xlWorksheet.Cells[row, 9] = data[9];

        //                string[] empresaData = line.Split('"');
        //                if (empresaData.Length >= 2)
        //                {
        //                    string[] empresaParts = empresaData[1].Split(';');
        //                    if (empresaParts.Length >= 1)
        //                    {
        //                        xlWorksheet.Cells[row, 10] = empresaParts[0];
        //                    }
        //                }

        //                for (int i = 11; i < 16; i++)
        //                {
        //                    xlWorksheet.Cells[row, i] = "";
        //                }
        //                foreach (string part in data.Skip(10))
        //                {
        //                    if (part.StartsWith("CONTENEDOR :"))
        //                    {
        //                        xlWorksheet.Cells[row, 11] = part.Remove(0, 13).Trim();
        //                    }
        //                    else if (part.StartsWith("TICKET DE PESAJE :"))
        //                    {
        //                        xlWorksheet.Cells[row, 12] = part.Remove(0, 19).Trim();
        //                    }
        //                    else if (part.StartsWith("PESO BRUTO :"))
        //                    {
        //                        xlWorksheet.Cells[row, 13] = part.Remove(0, 12).Trim();
        //                    }
        //                    else if (part.StartsWith("PESO TARA :"))
        //                    {
        //                        xlWorksheet.Cells[row, 14] = part.Remove(0, 11).Trim();
        //                    }
        //                    else if (part.StartsWith("PESO NETO :"))
        //                    {
        //                        xlWorksheet.Cells[row, 15] = part.Remove(0, 11).Trim();
        //                    }
        //                }
        //            }

        //            row++;
        //        }
        //    }

        //    //xlWorkbook.SaveAs(@"D:\LBB-GUIAS\output.xlsx");
        //    //xlWorkbook.Close();
        //    //xlApp.Quit();
        //    xlWorkbook.SaveAs(@"D:\LBB-GUIAS\output.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
        //               false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, false, false, Type.Missing, Type.Missing, Type.Missing);

        //    xlWorkbook.Close(false, Type.Missing, Type.Missing);
        //    xlApp.Quit();
        //}

        //private void ConvertTextToExcel()
        //{
        //    string[] filePaths = Directory.GetFiles(@"D:\LBB-GUIAS\", "*.txt");

        //    Excel.Application xlApp = new Excel.Application();
        //    Excel.Workbook xlWorkbook = xlApp.Workbooks.Add();
        //    Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
        //    Excel.Range xlRange = xlWorksheet.UsedRange;

        //    int row = 1;
        //    int[] rowsToExtract = new int[] { 1, 2, 5, 47, 50, 56, 59, 65, 68, 71, 89, 116, 120, 169, 172 };
        //    foreach (int rowToExtract in rowsToExtract)
        //        if (rowToExtract - 1 >= 0 && rowToExtract - 1 < filePaths.Length)
        //        {
        //            string filePath = filePaths[rowToExtract - 1];
        //            string[] lines = File.ReadAllLines(filePath);
        //            string line = lines.Last();
        //            string[] allData = line.Split(';');
        //            string columnValue = allData[allData.Length - 1];
        //            string[] data = lines[rowToExtract - 1].Split(';');
        //            for (int i = 0; i < data.Length; i++)
        //            {
        //                xlWorksheet.Cells[row, i + 1] = data[i];
        //            }
        //            xlWorksheet.Cells[row, data.Length + 1] = columnValue;
        //            row++;
        //        }

        //    xlWorkbook.SaveAs(@"D:\LBB-GUIAS\output.xlsx");
        //    xlWorkbook.Close();
        //    xlApp.Quit();
        //}

        private void ConvertTextToExcel()
        {
            string[] filePaths = Directory.GetFiles(@"D:\LBB-GUIAS\", "*.txt");

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Add();
            Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int row = 1;
            int[] rowsToExtract = new int[] { 1, 2, 5, 47, 50, 56, 59, 65, 68, 71, 89, 116, 120, 169, 172 };
            foreach (int rowToExtract in rowsToExtract)
                if (rowToExtract - 1 >= 0 && rowToExtract - 1 < filePaths.Length)
                {
                    string filePath = filePaths[rowToExtract - 1];
                    string[] lines = File.ReadAllLines(filePath);
                    for (int i = 0; i < lines.Length; i++)
                    {
                        if (lines[i].Contains("A;Serie;;") ||
                            lines[i].Contains("A;Correlativo;;") ||
                            lines[i].Contains("A;FchEmis;;") ||
                            lines[i].Contains("E;DescripcionAdicsunat;10;") ||
                            lines[i].Contains("E;DescripcionAdicsunat;11;T") ||
                            lines[i].Contains("E;DescripcionAdicsunat;13;") ||
                            lines[i].Contains("E;DescripcionAdicsunat;14;") ||
                            lines[i].Contains("E;DescripcionAdicsunat;16;") ||
                            lines[i].Contains("E;DescripcionAdicsunat;17;") ||
                            lines[i].Contains("E;DescripcionAdicsunat;18;") ||
                            lines[i].Contains("E;DescripcionAdicsunat;24;") ||
                            lines[i].Contains("E;DescripcionAdicsunat;31;") ||
                            lines[i].Contains("E;DescripcionAdicsunat;32;") ||
                            lines[i].Contains("G1;FechInicioTraslado;1;") ||
                            lines[i].Contains("G1;RazoTrans;1;"))
                        {
                            string[] allData = lines[i].Split(';');
                            string columnValue = allData[allData.Length - 1];
                            string[] data = lines[i + 1].Split(';');
                            for (int j = 0; j < data.Length; j++)
                            {
                                xlWorksheet.Cells[row, j + 1] = data[j];
                            }
                            xlWorksheet.Cells[row, data.Length + 1] = columnValue;
                            row++;
                        }
                    }
                }

            xlWorkbook.SaveAs(@"D:\LBB-GUIAS\output.xlsx");
            xlWorkbook.Close();
            xlApp.Quit();
        }


    }
}
