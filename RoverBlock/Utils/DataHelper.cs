using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;

namespace RoverBlock
{
    public static class DataHelper
    {
        public static List<Student> GetStudents(Dictionary<string, int> map, int grade)
        {
            string fileName = "Students" + grade + ".xls";
            List<Student> students = new List<Student>();

            List<Dictionary<string, string>> sheetData = SheetHelper.ReadSheet(fileName, map);
            foreach (Dictionary<string, string> entry in sheetData)
            {
                string NetworkID = entry["NetworkID"].ToLower();
                string LastName = entry["LastName"];
                string FirstName = entry["FirstName"];

                Student s = new Student(NetworkID, FirstName, LastName);

                students.Add(s);
            }

            return students;
        }

        public static void LoadStudentChoices(Dictionary<string, int> map, int grade, List<Student> students, List<Block> blocks)
        {
            List<string> blockIntersect = blocks.Where(x => x.aSlots != 0 || x.bSlots != 0).Select(x => x.Name).ToList();
            string fileName = "Choices" + grade + ".xls";
            List<Dictionary<string, string>> sheetData = SheetHelper.ReadSheet(fileName, map);
            foreach (Dictionary<string, string> entry in sheetData)
            {
                string NetworkID = entry["NetworkID"].ToLower().Replace("@roverkids.org", "");

                List<string> choices = new List<string>()
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

        public static void LockStudents(string fileName, Dictionary<string, int> map, List<Student> students)
        {
            List<Dictionary<string, string>> sheetData = SheetHelper.ReadSheet(fileName, map);
            foreach (Dictionary<string, string> entry in sheetData)
            {
                string NetworkID = entry["NetworkID"].ToLower();
                string BlockID = "RESERVED"; // entry["BlockID"];
                string BlockName = "RESERVED";
                string Day = entry["Day"];
                

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

        public static List<Block> GetBlocks(Dictionary<string, int> map, int grade, Random rnd)
        {
            string fileName = "Blocks" + grade + ".xls";
            List<Block> blocks = new List<Block>();

            List<Dictionary<string, string>> sheetData = SheetHelper.ReadSheet(fileName, map);
            foreach (Dictionary<string, string> entry in sheetData)
            {
                string BlockID = entry["Block ID"];
                string BlockName = entry["Block Name"].Trim();

                int aSlots = TryParse(entry["A Slots"]);
                int bSlots = TryParse(entry["B Slots"]);

                blocks.Add(new Block(BlockID, BlockName, aSlots, bSlots));
            }

            Shuffle(blocks, rnd);
            return blocks;
        }

        public static void RunLotteryA(Block b, List<Student> students, Random rnd)
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

        public static void RunLotteryB(Block b, List<Student> students, Random rnd)
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

        public static void AssignRemaining(List<Student> students, List<Block> blocks, Random rnd)
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

        public static int TryParse(string str)
        {
            return int.TryParse(str, out int result) ? result : 0;
        }

        public static void Shuffle<T>(List<T> list, Random rnd)
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

        public static T DeepCopy<T>(T obj)
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
