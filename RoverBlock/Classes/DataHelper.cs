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
            String fileName = "Choices" + grade + ".xls";
            List<Dictionary<String, String>> sheetData = sh.readSheet(fileName, map);
            foreach (Dictionary<String, String> entry in sheetData)
            {
                String NetworkID = entry["NetworkID"].ToLower().Replace("@roverkids.org", "");

                // TODO: use an intersect here. that would solve the duplicate issue and "promote" choices if the class does not exist
                List<String> choices = new List<String>()
                {
                    entry["Choice1"],
                    entry["Choice2"],
                    entry["Choice3"],
                    entry["Choice4"]
                }.Distinct().ToList();

                Student s = students.Where(x => x.NetworkID == NetworkID).FirstOrDefault();

                if(s == null)
                {
                    continue;
                }

                for(int i = 0; i < choices.Count; i++)
                {
                    String choice = choices[i];
                    
                    // try to match Name to ID
                    String RoverBlockID = blocks.Where(x => x.Name.ToLower() == choice.ToLower()).Select(x => x.ID).FirstOrDefault();

                    if(RoverBlockID == null)
                    {
                        RoverBlockID = choice;
                    }
                    else
                    {
                        RoverBlockID = "";
                    }

                    choices[i] = RoverBlockID;
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

                if (NetworkID != "")
                {
                    Student s = students.Where(x => x.NetworkID == NetworkID).FirstOrDefault();

                    if (s == null)
                    {
                        continue;
                    }

                    if (Day == "A")
                    {
                        s.A = new Block(BlockID, "", 0, 0); ;
                    }
                    else if (Day == "B")
                    {
                        s.B = new Block(BlockID, "", 0, 0); ;
                    }
                }
            }
        }

        public List<Block> getBlocks(String fileName, Dictionary<String, int> map)
        {
            List<Block> blocks = new List<Block>();

            List<Dictionary<String, String>> sheetData = sh.readSheet(fileName, map);
            foreach (Dictionary<String, String> entry in sheetData)
            {
                String BlockID = entry["Block ID"];
                String BlockName = entry["Block Name"];

                // int aSlots = tryParse(entry["A Slots"]);
                // int bSlots = tryParse(entry["B Slots"]);

                blocks.Add(new Block(BlockID, BlockName, 0, 0));
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
                pool.AddRange(Enumerable.Repeat(s, 3 - s.Choices.IndexOf(b.Name)));
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
