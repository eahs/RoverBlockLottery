using RoverBlock.Classes;
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
            DataHelper dh = new DataHelper();
            SheetHelper sh = new SheetHelper();
            ReconHelper rh = new ReconHelper();

            String output = "";

            Dictionary<String, int> blocksMap = new Dictionary<string, int>()
            {
                { "Block ID", 0 },
                { "Block Name", 1 },
                { "A Slots", 2 },
                { "B Slots", 3 }
            };

            Dictionary<String, int> studentsMap = new Dictionary<string, int>()
            {
                { "NetworkID", 0 },
                { "LastName", 1 },
                { "FirstName", 2 }
            };

            Dictionary<String, int> lockedStudentsMap = new Dictionary<String, int>()
            {
                { "Day", 4 },
                { "BlockID", 1 },
                { "NetworkID", 3 }
            };

            Dictionary<String, int> studentChoiceMap = new Dictionary<String, int>()
            {
                { "LastName", 1 },
                { "FirstName", 2 },
                { "NetworkID", 8 },
                { "Choice1", 4 },
                { "Choice2", 5 },
                { "Choice3", 6 },
                { "Choice4", 7 }
            };

            for (int i = 9; i < 12; i++)
            {
                Console.WriteLine("Grade " + i + " starting.");

                // initialize variables for scoring
                int score = int.MaxValue;
                List<Student> bestStudents = new List<Student>();
                List<Block> bestBlocks = new List<Block>();

                // load in data from sheets
                List<Block> baseBlocks = dh.getBlocks(blocksMap, i);
                List<Student> baseStudents = dh.getStudents(studentsMap, i);
                dh.lockStudents("LockedStudents.xls", lockedStudentsMap, baseStudents);
                dh.loadStudentChoices(studentChoiceMap, i, baseStudents, baseBlocks);

                // reconnaissance to make sense of certain data points
                rh.countChoices(studentChoiceMap, i, baseStudents);
                rh.noChoiceStudents(baseStudents, i);
                rh.wallOfShame(studentChoiceMap, i);

                for (int j = 0; j < 1000; j++)
                {
                    // deep copy of the above blocks and students
                    List<Block> blocks = baseBlocks.ConvertAll(x => dh.DeepCopy(x));
                    List<Student> students = baseStudents.ConvertAll(x => dh.DeepCopy(x));

                    dh.shuffle(blocks);

                    foreach (Block b in blocks)
                    {
                        /* if (b.aSlots != 0 && b.bSlots != 0 && students.Where(x => x.Choices != null && x.Choices.Contains(b.Name)).Count() == 0)
                        {
                            onsole.WriteLine("No grade " + i + " student voted for " + b.Name);
                        } */

                        dh.runLotteryA(b, students);
                        dh.runLotteryB(b, students);

                        // remove class from all students who did not win the lottery to promte their other choices
                        students.Where(x => x.Choices != null).Select(x => { x.Choices.Remove(b.Name); return x; }).ToList();
                    }



                    int newScore = students.Where(x => x.Choices != null && x.A == null).Count() + students.Where(x => x.Choices != null && x.B == null).Count();

                    if (newScore < score)
                    {
                        bestStudents = students.ConvertAll(x => dh.DeepCopy(x));
                        bestBlocks = blocks.ConvertAll(x => dh.DeepCopy(x));
                        score = newScore;
                    }
                }

                // dh.assignRemaining(bestStudents, bestBlocks);

                // output to XLS file
                sh.writeStudentsSheet(bestStudents, i);

                foreach (Block b in bestBlocks)
                {
                    int sumSlots = b.aSlots + b.bSlots;

                    if (sumSlots != 0)
                    {
                        output += sumSlots + " slots open in " + b.Name + "\n";
                    }
                }

                output += "A unscheduled: " + bestStudents.Where(x => x.Choices != null && x.A == null).Count() + "\n";
                output += "B unscheduled: " + bestStudents.Where(x => x.Choices != null && x.B == null).Count() + "\n";
                output += "Grade " + i + " score: " + score + " (lower is better)\n\n";
            }

            Console.WriteLine();

            output = output.Trim();

            File.WriteAllLines("../../Sheets/Output/ConsoleOutput.txt", output.Split('\n'));

            Console.WriteLine(output);

            Console.Write("\nScheduling completed! Press Enter to continue: ");
            Console.ReadLine();
        }
    }
}
