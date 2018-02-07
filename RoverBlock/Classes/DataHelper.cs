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
        private static SheetHelper sh = new SheetHelper();

        public List<Student> getStudents(Dictionary<String, int> map, int grade)
        {
            String fileName = "Students" + grade + ".xls";
            List<Student> students = new List<Student>();

            List<Dictionary<String, String>> sheetData = sh.readSheet(fileName, map);
            foreach (Dictionary<String, String> entry in sheetData)
            {
                String NetworkID = entry["NetworkID"].ToLower();
                String LastName = entry["LastName"];
                String FirstName = entry["FirstName"];

                Student s = new Student(NetworkID, FirstName, LastName);

                students.Add(s);
            }

            return students;
        }

        public void loadStudentChoices(Dictionary<String, int> map, int grade, List<Student> students, List<Block> blocks)
        {
            List<String> blockIntersect = blocks.Where(x => x.aSlots != 0 || x.bSlots != 0).Select(x => x.Name).ToList();
            String fileName = "Choices" + grade + ".xls";
            List<Dictionary<String, String>> sheetData = sh.readSheet(fileName, map);
            foreach (Dictionary<String, String> entry in sheetData)
            {
                String NetworkID = entry["NetworkID"].ToLower().Replace("@roverkids.org", "");

                List<String> choices = new List<String>()
                {
                    entry["Choice1"],
                    entry["Choice2"],
                    entry["Choice3"],
                    entry["Choice4"]
                }.Intersect(blockIntersect).ToList();

                Student s = students.Where(x => x.NetworkID == NetworkID).FirstOrDefault();

                if (s == null)
                {
                    continue;
                }

                s.Choices = choices;
            }
        }

        public void lockStudents(String fileName, Dictionary<String, int> map, List<Student> students)
        {
            List<Dictionary<String, String>> sheetData = sh.readSheet(fileName, map);
            foreach (Dictionary<String, String> entry in sheetData)
            {
                String NetworkID = entry["NetworkID"].ToLower();
                String BlockID = entry["BlockID"];
                String Day = entry["Day"];

                // TODO: remove this when BlockIDs are in the sheet
                BlockID = "RESERVED";

                if (NetworkID != "")
                {
                    Student s = students.Where(x => x.NetworkID == NetworkID).FirstOrDefault();

                    if (s == null)
                    {
                        continue;
                    }

                    if (Day == "A")
                    {
                        s.A = new Block(BlockID, "RESERVED", 0, 0); ;
                    }
                    else if (Day == "B")
                    {
                        s.B = new Block(BlockID, "RESERVED", 0, 0); ;
                    }
                }
            }
        }

        public List<Block> getBlocks(Dictionary<String, int> map, int grade)
        {
            String fileName = "Blocks" + grade + ".xls";
            List<Block> blocks = new List<Block>();

            List<Dictionary<String, String>> sheetData = sh.readSheet(fileName, map);
            foreach (Dictionary<String, String> entry in sheetData)
            {
                String BlockID = entry["Block ID"];
                String BlockName = entry["Block Name"].Trim();

                int aSlots = tryParse(entry["A Slots"]);
                int bSlots = tryParse(entry["B Slots"]);

                blocks.Add(new Block(BlockID, BlockName, aSlots, bSlots));
            }

            shuffle(blocks);
            return blocks;
        }

        public void runLotteryA(Block b, List<Student> students)
        {
            List<Student> interestedStudents = students.Where(x => x.Choices != null && x.A == null && x.Choices.Contains(b.Name)).ToList();
            List<Student> pool = new List<Student>();

            foreach (Student s in interestedStudents)
            {
                pool.AddRange(Enumerable.Repeat(s, 4 - s.Choices.IndexOf(b.Name)));
            }

            interestedStudents = null;
            GC.Collect();

            while (b.aSlots > 0)
            {
                if (pool.Count == 0)
                {
                    break;
                }

                Student winner = pool[rnd.Next(pool.Count)];
                winner.Choices.Remove(b.Name);
                winner.A = b;
                pool.RemoveAll(x => x.NetworkID == winner.NetworkID);

                b.aSlots--;
            }
        }

        public void runLotteryB(Block b, List<Student> students)
        {
            List<Student> interestedStudents = students.Where(x => x.Choices != null && x.B == null && x.Choices.Contains(b.Name)).ToList();
            List<Student> pool = new List<Student>();

            foreach (Student s in interestedStudents)
            {
                pool.AddRange(Enumerable.Repeat(s, 4 - s.Choices.IndexOf(b.Name)));
            }

            interestedStudents = null;
            GC.Collect();

            while (b.bSlots > 0)
            {
                if (pool.Count == 0)
                {
                    break;
                }

                Student winner = pool[rnd.Next(pool.Count)];
                winner.Choices.Remove(b.Name);
                winner.B = b;
                pool.RemoveAll(x => x.NetworkID == winner.NetworkID);

                b.bSlots--;
            }
        }

        public int tryParse(String str)
        {
            return int.TryParse(str, out int result) ? result : 0;
        }

        public void shuffle<T>(List<T> list)
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
