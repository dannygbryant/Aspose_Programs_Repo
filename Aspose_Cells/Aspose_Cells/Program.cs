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
        static void Main(string[] args)
        {
            string path = "";
            Console.WriteLine("Enter File Path");
            path = Console.ReadLine();
            Console.WriteLine(path);
            CreateWorkbook(path);
            Console.ReadKey();
            //CreateWorkbook(path);
        }

        static void CreateWorkbook(string path)
        {
            Workbook wb = new Workbook();
            wb.Save(path);
        }
    }
}
