using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;

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
                }.Distinct().ToList(); // .Intersect(blockIntersect)

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
                String BlockID = "RESERVED"; // entry["BlockID"];
                String BlockName = "RESERVED";
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
                        s.A = new Block(BlockID, BlockName, 0, 0); ;
                    }
                    else if (Day == "B")
                    {
                        s.B = new Block(BlockID, BlockName, 0, 0); ;
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

        public void assignRemaining(List<Student> students, List<Block> blocks)
        {
            List<Block> aBlocks = blocks.Where(x => x.aSlots != 0).ToList();
            List<Block> bBlocks = blocks.Where(x => x.bSlots != 0).ToList();

            foreach (Student s in students.Where(x => x.Choices != null && (x.A == null || x.B == null)))
            {
                if(s.A == null)
                {
                    if (aBlocks.Count != 0)
                    {
                        Block b = aBlocks[rnd.Next(aBlocks.Count)];
                        s.A = b;
                        b.aSlots--;

                        if (b.aSlots == 0)
                        {
                            aBlocks.Remove(b);
                        }
                    }
                }

                if (s.B == null)
                {
                    if (bBlocks.Count != 0)
                    {
                        Block b = bBlocks[rnd.Next(bBlocks.Count)];
                        s.B = b;
                        b.bSlots--;

                        if (b.bSlots == 0)
                        {
                            bBlocks.Remove(b);
                        }
                    }
                }
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

        public T DeepCopy<T>(T obj)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                BinaryFormatter formatter = new BinaryFormatter();
                formatter.Serialize(stream, obj);
                stream.Position = 0;

                return (T)formatter.Deserialize(stream);
            }
        }
    }
}
