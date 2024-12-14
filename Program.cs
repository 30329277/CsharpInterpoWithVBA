using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CsharpInterpoWithVBA
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // 创建Excel应用程序对象
            Excel.Application excelApp = new Excel.Application();
            // 打开Excel文件
            Excel.Workbook workbook = excelApp.Workbooks.Open(@"D:\DotNetProject\forCsharpTest.xlsm");

            // 调用宏
            excelApp.Run("Module1.Test1");

            // 获取第一个工作表
            Excel.Worksheet worksheet = workbook.Sheets["Sheet1"];
            // 读取单元格A1的值
            string cellValue = worksheet.Range["A1"].Value.ToString();

            // 打印单元格A1的值到控制台
            Console.WriteLine("Value in Sheet1 A1: " + cellValue);

            // 关闭工作簿
            workbook.Close(false);
            // 退出Excel应用程序
            excelApp.Quit();

            // 释放COM对象
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            Console.WriteLine("Macro executed successfully.");
            Console.ReadKey();
        }
    }
}
