using RoverBlock.Utils;
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
                string LastName = entry.ContainsKey("LastName") ? entry["LastName"] : "";
                string FirstName = entry.ContainsKey("FirstName") ? entry["FirstName"] : "";

                Student s = new Student
                {
                    NetworkID = NetworkID,
                    FirstName = FirstName,
                    LastName = LastName
                };

                students.Add(s);
            }

            return students;
        }

        public static void LoadStudentChoices(Dictionary<string, int> map, int grade, List<Student> students, List<Block> blocks = null)
        {
            List<string> blockIntersect = null;

            if(blocks != null)
            {
                blockIntersect = blocks.Where(b => b.Slots != 0).Select(x => x.Name).ToList();
            }

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
                }.Distinct().ToList();

                if(blocks != null)
                {
                    choices = choices.Intersect(blockIntersect).ToList();
                }

                choices = choices.PadLeft(4, null).ToList();

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
                
                if (NetworkID != "")
                {
                    Student s = students.Where(x => x.NetworkID == NetworkID).FirstOrDefault();

                    if (s == null)
                    {
                        continue;
                    }

                    s.RoverBlock = new Block(BlockID, BlockName, 0);
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

                int slots = TryParse(entry["Slots"]);

                blocks.Add(new Block(BlockID, BlockName, slots));
            }

            Shuffle(blocks, rnd);
            return blocks;
        }

        public static void RunLottery(Block b, List<Student> students, Random rnd)
        {
            List<Student> interestedStudents = students.Where(s => s.Choices != null && s.RoverBlock == null && s.Choices.Contains(b.Name)).ToList();
            List<Student> pool = new List<Student>();

            foreach (Student s in interestedStudents)
            {
                pool.AddRange(Enumerable.Repeat(s, 4 - s.Choices.IndexOf(b.Name)));
            }

            interestedStudents = null;

            while (b.Slots > 0)
            {
                if (pool.Count == 0)
                {
                    break;
                }

                Student winner = pool[rnd.Next(pool.Count)];

                int choiceIdx = winner.Choices.IndexOf(b.Name);

                // assign a score based on the priority of the student's choice
                winner.LotteryScore = 4 - choiceIdx;

                winner.Choices[choiceIdx] = null;
                winner.RoverBlock = b;
                pool.RemoveAll(x => x.NetworkID == winner.NetworkID);

                b.Slots--;
            }
        }

        public static void AssignRemaining(List<Student> students, List<Block> blocks, Random rnd)
        {
            List<Block> OpenBlocks = blocks.Where(b => b.Slots != 0).ToList();

            foreach (Student s in students.Where(s => s.Choices != null && (s.RoverBlock == null)))
            {
                if(s.RoverBlock == null)
                {
                    if (OpenBlocks.Count != 0)
                    {
                        Block b = OpenBlocks[rnd.Next(OpenBlocks.Count)];
                        s.RoverBlock = b;

                        if (b.Slots == 0)
                        {
                            OpenBlocks.Remove(b);
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
