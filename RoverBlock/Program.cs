using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RoverBlock
{
    class Program
    {
        static void Main(string[] args)
        {
            Random rnd = new Random();
            string output = "";

            Dictionary<string, int> blocksMap = new Dictionary<string, int>()
            {
                { "Block ID", 0 },
                { "Block Name", 1 },
                { "Slots", 2 },
            };

            Dictionary<string, int> studentsMap = new Dictionary<string, int>()
            {
                { "NetworkID", 0 },
                { "LastName", 1 },
                { "FirstName", 2 }
            };

            Dictionary<string, int> lockedStudentsMap = new Dictionary<string, int>()
            {
                { "Day", 4 },
                { "BlockID", 1 },
                { "NetworkID", 3 }
            };

            Dictionary<string, int> studentChoiceMap = new Dictionary<string, int>()
            {
                { "LastName", 0 },
                { "FirstName", 1 },
                { "NetworkID", 7 },
                { "Choice1", 3 },
                { "Choice2", 4 },
                { "Choice3", 5 },
                { "Choice4", 6 }
            };

            for (int i = 9; i < 12; i++)
            {
                Console.WriteLine("Grade " + i + " starting.");

                // initialize variables for scoring
                double efficiency = 0;
                int totalStudents = 0;
                int score = int.MinValue;
                List<Student> bestStudents = new List<Student>();
                List<Block> bestBlocks = new List<Block>();

                // load in data from sheets
                List<Block> baseBlocks = DataHelper.GetBlocks(blocksMap, i, rnd);
                List<Student> baseStudents = DataHelper.GetStudents(studentsMap, i);
                // DataHelper.LockStudents("LockedStudents.xls", lockedStudentsMap, baseStudents);
                DataHelper.LoadStudentChoices(studentChoiceMap, i, baseStudents, baseBlocks);

                // reconnaissance to make sense of certain data points
                ReconHelper.CountChoices(studentChoiceMap, i, baseStudents);
                ReconHelper.NoChoiceStudents(baseStudents, i);
                ReconHelper.WallOfShame(studentChoiceMap, i);

                Console.CursorVisible = false;
                for (int j = 0; j < 1000; j++)
                {
                    Console.Write("\rTrying combination #{0}  |  Best Score: {1}", j, score.ToString().PadLeft(6));

                    // deep copy of the above blocks and students
                    List<Block> blocks = baseBlocks.ConvertAll(x => DataHelper.DeepCopy(x));
                    List<Student> students = baseStudents.ConvertAll(x => DataHelper.DeepCopy(x));

                    DataHelper.Shuffle(blocks, rnd);

                    foreach (Block b in blocks)
                    {
                        /* if (b.aSlots != 0 && b.bSlots != 0 && students.Where(x => x.Choices != null && x.Choices.Contains(b.Name)).Count() == 0)
                        {
                            onsole.WriteLine("No grade " + i + " student voted for " + b.Name);
                        } */

                        DataHelper.RunLottery(b, students, rnd);

                        // remove class from all students who did not win the lottery to promte their other choices
                        students.Where(x => x.Choices != null).Select(s => { s.Choices.Remove(b.Name); return s; }).ToList();
                    }

                    int newScore = students.Select(x => x.LotteryScore).Sum();

                    if (newScore > score)
                    {
                        bestStudents = students.ConvertAll(x => DataHelper.DeepCopy(x));
                        bestBlocks = blocks.ConvertAll(x => DataHelper.DeepCopy(x));
                        score = newScore;
                        totalStudents = students.Where(x => x.LotteryScore > 0).Count();
                    }

                    efficiency = (double)score / (Math.Pow(2, 4) * totalStudents) * 100;
                    efficiency = Math.Round(efficiency, 3);

                    Console.Write(" |  Latest Score: {0}  |  Efficiency: {1}%    ", newScore.ToString().PadLeft(6), efficiency);
                }
                Console.CursorVisible = true;

                output += "Grade " + i + " unscheduled (Initial): " + bestStudents.Where(x => x.RoverBlock == null).Count() + "\n";

                DataHelper.AssignRemaining(bestStudents, bestBlocks, rnd);

                // output to XLS file
                SheetHelper.WriteStudentsSheet(bestStudents, i);

                foreach (Block b in bestBlocks)
                {
                    if (b.Slots != 0)
                    {
                        output += b.Slots + " slots open in " + b.Name + "\n";
                    }
                }

                output += "Score: " + score + "\n";
                output += "Grade " + i + " unscheduled (Final): " + bestStudents.Where(x => x.RoverBlock == null).Count() + "\n\n";
            }

            Console.WriteLine("\n");

            output = output.Trim();

            File.WriteAllLines("../../Sheets/Output/ConsoleOutput.txt", output.Split('\n'));

            Console.WriteLine(output);

            Console.Write("\nScheduling completed! Press Enter to continue: ");
            Console.ReadLine();
        }
    }
}
