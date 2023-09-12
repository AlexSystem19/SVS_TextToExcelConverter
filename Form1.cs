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

        private void ConvertTextToExcel()
        {
            string[] filePaths = Directory.GetFiles(@"D:\LBB-GUIAS\", "*.txt");

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Add();
            Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int row = 1;
            foreach (string filePath in filePaths)
            {
                string[] lines = File.ReadAllLines(filePath);
                foreach (string line in lines)
                {
                    string[] data = line.Split(',');
                    for (int i = 0; i < data.Length; i++)
                    {
                        xlWorksheet.Cells[row, i + 1] = data[i];
                    }
                    row++;
                }
            }

            xlWorkbook.SaveAs(@"D:\LBB-GUIAS\output.xlsx");
            xlWorkbook.Close();
            xlApp.Quit();
        }
    }
}
