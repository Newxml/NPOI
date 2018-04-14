using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI
{
    public class Program
    {
        public static bool creExcel()
        {
            bool cre = false;
            //创建文件  
            XSSFWorkbook wk = new XSSFWorkbook();

            //创建Excel工作表  
            var sheet = wk.CreateSheet("第一个Sheet");

            //创建单元格  
            for (int i = 0; i < 100; i++)
            {
                int n = (int)Math.Floor(i / 10d);
                if (n == 0)
                {
                    var row = sheet.CreateRow(n);
                    var cell = row.CreateCell(i);
                    cell.SetCellValue(n + "+" + (i+1));
                }
                else
                {
                    int m = i % (n * 10);
                    var row = sheet.CreateRow(n);
                    var cell = row.CreateCell(m);
                    cell.SetCellValue(n + "+" + (i+1));
                }

            }
            //保存
            FileStream file = new FileStream(@"E:\test.xlsx", FileMode.Create);
            wk.Write(file);
            file.Close();
            if (!string.IsNullOrEmpty(file.Name))
            {
                cre = true;
            }
            return cre;
        }

        static void Main(string[] args)
        {
            bool a = creExcel();


            if (a)
            {
                Console.Write("新建Excel成功");
            }
            else
            {
                Console.Write("创建Excel失败");
            }
            Console.ReadLine();
        }
    }
}
