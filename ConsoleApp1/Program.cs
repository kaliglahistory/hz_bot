using CsvHelper;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Util;
using DocumentFormat.OpenXml.Spreadsheet;
using Aspose.Cells;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Bibliography;
using System.Text;
using NPOI.SS.Formula.Functions;
// Открытие существующей рабочей книги
IWorkbook workbook;
using (FileStream fileStream = new FileStream("C:\\Users\\user\\Desktop\\программы\\asc\\Шаблон (9).xlsx", FileMode.Open, FileAccess.Read))
{
    workbook = new XSSFWorkbook(fileStream);
}
string filePaths = $"C:\\Users\\user\\Desktop\\программы\\asc\\База{1+35}.csv";
string filePaths2 = @"C:\\Users\user\Desktop\программы\asc\";
// Получение листа
ISheet sheet = workbook.GetSheetAt(0);
int a = sheet.LastRowNum;
string[] strings = new string[a+1];
StringBuilder csv = new StringBuilder( );
File.Create(filePaths);

FileStream f = new FileStream(filePaths, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    f.WriteByte(123);
f.Position = 0;
//StreamWriter file = new StreamWriter(filePaths);
//file.Close();zz
        f.Close();
StreamReader sr = new StreamReader(filePaths, System.Text.Encoding.Default);
        f.Close();   
File.Create(filePaths).Close();
f.Close();
for (int j = 0; j <= a; j++)
{

    int i = j;
    // Чтение данных из ячейки
    IRow row = sheet.GetRow(j);
    string cellValue = row.GetCell(0).StringCellValue;
    strings[j] = cellValue;
    
        csv.Append("1,КИЗ" + strings[j] + ","+ '"');
            File.AppendAllText(filePaths, csv.ToString(),Encoding.Default);
    
         ///strings[i] = cellValue;
        // Вывод данных ячейки
    Console.WriteLine(cellValue);
   
}


//string[] stringsa = new string[a];
//for (int s = 0; s < a; s++)
//{
//    csv.AppendLine(stringsa[s] + ",");
//    File.AppendAllText(filePaths, csv.ToString());
//}

//// Загрузите исходный файл Excel




