using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RoverBlock.Classes
{
    class DataHelper
    {
        private static Random rnd = new Random();

        public List<Student> getLockedStudents(String fileName, Dictionary<String, int> map)
        {
            List<LockedStudent> lockedStudents = new List<LockedStudent>();

            List<Dictionary<String, String>> sheetData = readSheet(fileName, map);
            foreach (Dictionary<String, String> entry in sheetData)
            {
                lockedStudents.Add(new LockedStudent(entry["NetworkID"], entry["Day"]));
            }

            Block reservedBlock = new Block("RESERVED", 0, 0);

            List<Student> output = (
                from ls in lockedStudents
                join lsA in lockedStudents
                    on new { a = ls.NetworkID, b = "A" } equals new { a = lsA.NetworkID, b = lsA.Day } into joinedA
                from lsA in joinedA.DefaultIfEmpty()
                join lsB in lockedStudents
                    on new { a = ls.NetworkID, b = "B" } equals new { a = lsB.NetworkID, b = lsB.Day } into joinedB
                from lsB in joinedB.DefaultIfEmpty()
                select new Student(ls.NetworkID, lsB == null ? null : reservedBlock, lsA == null ? null : reservedBlock)
            ).GroupBy(x => x.NetworkID).Select(x => x.FirstOrDefault()).ToList();

            return output;
        }

        public List<Student> getStudents(String fileName, Dictionary<String, int> map, List<Student> lockedStudents)
        {
            List<Student> students = new List<Student>();

            List<Dictionary<String, String>> sheetData = readSheet(fileName, map);
            foreach (Dictionary<String, String> entry in sheetData)
            {
                String NetworkID = entry["NetworkID"];
                List<String> choices = new List<String>()
                {
                    entry["Choice 1"],
                    entry["Choice 2"],
                    entry["Choice 3"]
                }.Distinct().ToList();

                // TODO: use an intersect here. that would solve the duplicate issue and "promote" choices if the class does not exist

                // reuse Student object from locked students to preserve RESERVED classes
                Student s = lockedStudents.Where(x => x.NetworkID == NetworkID).SingleOrDefault();

                // if no student matched, create a new one
                if (s == null)
                {
                    s = new Student(NetworkID, null, null);
                }

                s.Choices = choices;

                students.Add(s);
            }

            return students;
        }

        public List<Block> getBlocks(String fileName, Dictionary<String, int> map)
        {
            List<Block> blocks = new List<Block>();

            List<Dictionary<String, String>> sheetData = readSheet(fileName, map);
            foreach (Dictionary<String, String> entry in sheetData)
            {
                String BlockName = entry["Class Name"];
                int aSlots = tryParse(entry["A Slots"]);
                int bSlots = tryParse(entry["B Slots"]);

                blocks.Add(new Block(BlockName, aSlots, bSlots));
            }

            shuffle(blocks);
            return blocks;
        }

        public void runLotteryA(Block b, List<Student> students)
        {
            List<Student> interestedStudents = students.Where(x => x.A == null && x.Choices.Contains(b.Name)).ToList();
            List<Student> pool = new List<Student>();

            foreach (Student s in interestedStudents)
            {
                pool.AddRange(Enumerable.Repeat(s, 4 - s.Choices.IndexOf(b.Name)));
            }

            interestedStudents = null;
            GC.Collect();

            for (int i = 0; i < b.aSlots; i++)
            {
                if (pool.Count == 0)
                {
                    break;
                }

                Student winner = pool[rnd.Next(pool.Count)];
                winner.Choices.Remove(b.Name);
                winner.A = b;
                pool.RemoveAll(x => x.NetworkID == winner.NetworkID);
            }
        }

        public void runLotteryB(Block b, List<Student> students)
        {
            List<Student> interestedStudents = students.Where(x => x.B == null && x.Choices.Contains(b.Name)).ToList();
            List<Student> pool = new List<Student>();

            foreach (Student s in interestedStudents)
            {
                pool.AddRange(Enumerable.Repeat(s, 4 - s.Choices.IndexOf(b.Name)));
            }

            interestedStudents = null;
            GC.Collect();

            for (int i = 0; i < b.bSlots; i++)
            {
                if (pool.Count == 0)
                {
                    break;
                }

                Student winner = pool[rnd.Next(pool.Count)];
                winner.Choices.Remove(b.Name);
                winner.B = b;
                pool.RemoveAll(x => x.NetworkID == winner.NetworkID);
            }
        }

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

        public void writeStudentsSheet(List<Student> students)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            IRow header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("Network ID");
            header.CreateCell(1).SetCellValue("A Day Class");
            header.CreateCell(2).SetCellValue("B Day Class");
            for (int i = 0; i < students.Count; i++)
            {
                Student student = students[i];
                IRow row = sheet.CreateRow(i + 1);
                row.CreateCell(0).SetCellValue(student.NetworkID);
                row.CreateCell(1).SetCellValue(student.A == null ? "" : student.A.Name);
                row.CreateCell(2).SetCellValue(student.B == null ? "" : student.B.Name);
            }
            for (int i = 0; i < 3; i++)
            {
                sheet.AutoSizeColumn(i);
            }
            using (var file = File.Create("../../Sheets/Output.xls"))
            {
                workbook.Write(file);
            }
        }


        public int tryParse(String str)
        {
            return int.TryParse(str, out int result) ? result : 0;
        }

        public static void shuffle<T>(List<T> list)
        {
            int n = list.Count;
            while (n > 1)
            {
                n--;
                int k = rnd.Next(n + 1);
                T value = list[k];
                list[k] = list[n];
                list[n] = value;
            }
        }
    }
}
