using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace RoverBlock
{
    public static class SheetHelper
    {
        public static List<Dictionary<string, string>> ReadSheet(string fileName, Dictionary<string, int> map)
        {
            List<Dictionary<string, string>> output = new List<Dictionary<string, string>>();

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

                    Dictionary<string, string> dict = new Dictionary<string, string>();

                    // The cells list doesn't necessary match up with what you see in excel, particularly if columns are left blank
                    Dictionary<int, int> columnIndexKeys = new Dictionary<int, int>();
                    int i = 0;
                    foreach (ICell cell in cells)
                    {
                        columnIndexKeys.Add(cell.ColumnIndex, i);
                        i++;
                    }

                    foreach (KeyValuePair<string, int> entry in map)
                    {
                        // avoid index out of range when reading spreadsheet data
                        if(columnIndexKeys.ContainsKey(entry.Value))
                        {
                            Debug.Assert(cells[columnIndexKeys[entry.Value]].ColumnIndex == entry.Value, "Column mismatch - did you leave a column blank?");
                            
                            dict.Add(entry.Key, cells[columnIndexKeys[entry.Value]].ToString());
                        }
                        else
                        {
                            dict.Add(entry.Key, "");
                        }
                    }

                    output.Add(dict);
                }
            }

            return output;
        }

        // TODO: add teacher name and students first/last name to row
        public static void WriteStudentsSheet(List<Student> students, int grade)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            IRow header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("Network ID");
            header.CreateCell(1).SetCellValue("Class ID");
            header.CreateCell(2).SetCellValue("Class Name");
            for (int i = 0; i < students.Count; i++)
            {
                Student student = students[i];
                IRow row = sheet.CreateRow(i + 1);
                row.CreateCell(0).SetCellValue(student.NetworkID);
                row.CreateCell(1).SetCellValue(student.RoverBlock == null ? "" : student.RoverBlock.ID);
                row.CreateCell(2).SetCellValue(student.RoverBlock == null ? "" : student.RoverBlock.Name);
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

        public static void WriteChoicesSheet(List<Dictionary<string, int>> choiceCounts, int grade)
        {
            List<string> classes = choiceCounts.SelectMany(x => x.Keys).Distinct().ToList();

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
                string className = classes[i];
                IRow row = sheet.CreateRow(i + 1);
                row.CreateCell(0).SetCellValue(className);
                for (int j = 0; j < 4; j++)
                {
                    if(choiceCounts[j].ContainsKey(className))
                    {
                        row.CreateCell(j + 1).SetCellValue(choiceCounts[j][className]);
                    }
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

        public static void WriteNoChoicesSheet(List<Student> noChoices, int grade)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            IRow header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("Network ID");
            header.CreateCell(1).SetCellValue("First Name");
            header.CreateCell(2).SetCellValue("Last Name");

            for (int i = 0; i < noChoices.Count; i++)
            {
                IRow row = sheet.CreateRow(i + 1);
                row.CreateCell(0).SetCellValue(noChoices[i].NetworkID);
                row.CreateCell(1).SetCellValue(noChoices[i].FirstName);
                row.CreateCell(2).SetCellValue(noChoices[i].LastName);
            }
            for (int i = 0; i < 3; i++)
            {
                sheet.AutoSizeColumn(i);
            }
            using (var file = File.Create("../../Sheets/Output/NoChoices" + grade + ".xls"))
            {
                workbook.Write(file);
            }
        }

        public static void WriteWallOfShame(List<Student> students, int grade)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            IRow header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("Network ID");
            header.CreateCell(1).SetCellValue("First Name");
            header.CreateCell(2).SetCellValue("Last Name");
            header.CreateCell(3).SetCellValue("Choices");

            for (int i = 0; i < students.Count; i++)
            {
                IRow row = sheet.CreateRow(i + 1);
                row.CreateCell(0).SetCellValue(students[i].NetworkID);
                row.CreateCell(1).SetCellValue(students[i].FirstName);
                row.CreateCell(2).SetCellValue(students[i].LastName);
                row.CreateCell(3).SetCellValue(string.Join(", ", students[i].Choices));
            }
            for (int i = 0; i < 4; i++)
            {
                sheet.AutoSizeColumn(i);
            }
            using (var file = File.Create("../../Sheets/Output/WallOfShame" + grade + ".xls"))
            {
                workbook.Write(file);
            }
        }
    }
}
