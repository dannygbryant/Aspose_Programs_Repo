using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;

namespace Aspose_Cells
{
    class Program
    {
        static string path;
        static void Main(string[] args)
        {
            path = "";
            Console.WriteLine("Enter File Path");
            path = Console.ReadLine();
            Console.WriteLine(path);
            CreateWorkbook(path);

            // Write Data to Cell
            Console.WriteLine("Enter Cell");
            string cell = Console.ReadLine();
            Console.WriteLine("Enter Text");
            string text = Console.ReadLine();
            Console.WriteLine("Enter workbook path");
            Workbook book = new Workbook( path );
            WriteDataToCell(text,book,cell);

            Console.ReadKey();
            
        }

        static void CreateWorkbook(string path)
        {
            Workbook wb = new Workbook();
            wb.Save(path);
        }

        static void WriteDataToCell(string text, Workbook wb, string cell)
        {
            wb.Worksheets[0].Cells[cell].Value = text;
            wb.Save(path);
            
            //Cell cell = new Cell();
        }
    }
}
