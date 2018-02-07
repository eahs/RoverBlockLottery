using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RoverBlock.Classes
{
    class SheetHelper
    {
        public List<Dictionary<String, String>> readSheet(String fileName, Dictionary<String, int> map)
        {
            List<Dictionary<String, String>> output = new List<Dictionary<String, String>>();

            HSSFWorkbook workbook;
            using (FileStream file = new FileStream("../../Sheets/" + fileName, FileMode.Open, FileAccess.Read))
            {
                workbook = new HSSFWorkbook(file);
            }

            ISheet sheet = workbook.GetSheetAt(0);
            for (int row = 1; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null)
                {
                    List<ICell> cells = sheet.GetRow(row).Cells;

                    if(cells.Count == 0)
                    {
                        break;
                    }

                    Dictionary<String, String> dict = new Dictionary<String, String>();

                    foreach (KeyValuePair<String, int> entry in map)
                    {
                        dict.Add(entry.Key, cells[entry.Value].ToString());
                    }

                    output.Add(dict);
                }
            }

            return output;
        }

        public void writeStudentsSheet(List<Student> students, int grade)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            IRow header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("Network ID");
            header.CreateCell(1).SetCellValue("A Day Class ID");
            header.CreateCell(2).SetCellValue("B Day Class ID");
            header.CreateCell(3).SetCellValue("A Day Class Name");
            header.CreateCell(4).SetCellValue("B Day Class Name");
            for (int i = 0; i < students.Count; i++)
            {
                Student student = students[i];
                IRow row = sheet.CreateRow(i + 1);
                row.CreateCell(0).SetCellValue(student.NetworkID);
                row.CreateCell(1).SetCellValue(student.A == null ? "" : student.A.ID);
                row.CreateCell(2).SetCellValue(student.B == null ? "" : student.B.ID);
                row.CreateCell(3).SetCellValue(student.A == null ? "" : student.A.Name);
                row.CreateCell(4).SetCellValue(student.B == null ? "" : student.B.Name);
            }
            for (int i = 0; i < 5; i++)
            {
                sheet.AutoSizeColumn(i);
            }
            using (var file = File.Create("../../Sheets/Output/Output" + grade + ".xls"))
            {
                workbook.Write(file);
            }
        }

        public void writeChoicesSheet(List<Dictionary<String, int>> choiceCounts, int grade)
        {
            List<String> classes = choiceCounts.SelectMany(x => x.Keys).Distinct().ToList();

            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            IRow header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("Class Name");
            header.CreateCell(1).SetCellValue("1st Choice Count");
            header.CreateCell(2).SetCellValue("2nd Choice Count");
            header.CreateCell(3).SetCellValue("3rd Choice Count");
            header.CreateCell(4).SetCellValue("4th Choice Count");

            for (int i = 0; i < classes.Count; i++)
            {
                String className = classes[i];
                IRow row = sheet.CreateRow(i + 1);
                row.CreateCell(0).SetCellValue(className);
                for (int j = 0; j < 4; j++)
                {
                    row.CreateCell(j + 1).SetCellValue(choiceCounts[j][className]);
                }
            }
            for (int i = 0; i < 5; i++)
            {
                sheet.AutoSizeColumn(i);
            }
            using (var file = File.Create("../../Sheets/Output/ChoiceCount" + grade + ".xls"))
            {
                workbook.Write(file);
            }
        }
    }
}
