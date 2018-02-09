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

            // TODO: random scheduling if they didn't get their choices

            for (int i = 9; i < 12; i++)
            {
                Console.WriteLine("Grade " + i + " starting.");

                int score = int.MaxValue;
                List<Student> bestStudents = new List<Student>();
                List<Block> bestBlocks = new List<Block>();

                // wall of shame for students who picked the same choice more than once
                rh.wallOfShame(studentChoiceMap, i);

                // TODO: don't run these operations on each iteration
                for (int j = 0; j < 1000; j++) {

                    // load blocks from master block list
                    List<Block> blocks = dh.getBlocks(blocksMap, i);

                    // load students from master directed study lists
                    List<Student> students = dh.getStudents(studentsMap, i);

                    // lock students into specific rover blocks
                    dh.lockStudents("LockedStudents.xls", lockedStudentsMap, students);

                    // load the students' choices into a list and remove duplicates
                    dh.loadStudentChoices(studentChoiceMap, i, students, blocks);

                    // compute interest in specific rover blocks for grades 9 through 11
                    rh.countChoices(studentChoiceMap, i, students);

                    // generate a list of students that did not fill out the google form
                    rh.noChoiceStudents(students, i);



                    // run lotteries in ascending order of popularity
                    // blocks = blocks.OrderBy(x => students.Where(y => y.Choices != null && (y.A == null || y.B == null) && y.Choices.Contains(x.Name)).Count()).ToList();

                    // Console.WriteLine("Grade " + i + " lotteries, running from least to most popular: ");
                    // Console.WriteLine(String.Join(", ", blocks.Select(x => x.Name)) + "\n");



                    foreach (Block b in blocks)
                    {
                        if (b.aSlots != 0 && b.bSlots != 0 && students.Where(x => x.Choices != null && x.Choices.Contains(b.Name)).Count() == 0)
                        {
                            // Console.WriteLine("No grade " + i + " student voted for " + b.Name);
                        }

                        dh.runLotteryA(b, students);
                        dh.runLotteryB(b, students);

                        // remove class from all students who did not win the lottery to promte their other choices
                        students.Where(x => x.Choices != null).Select(x => { x.Choices.Remove(b.Name); return x; }).ToList();
                    }



                    int newScore = students.Where(x => x.Choices != null && x.A == null).Count() + students.Where(x => x.Choices != null && x.B == null).Count();

                    if(newScore < score)
                    {
                        bestStudents = students.ToList();
                        bestBlocks = blocks.ToList();
                        score = newScore;
                    }
                }

                // dh.assignRemaining(bestStudents, bestBlocks);

                // output to XLS file
                sh.writeStudentsSheet(bestStudents, i);

                foreach (Block b in bestBlocks)
                {
                    int sumSlots = b.aSlots + b.bSlots;

                    if(sumSlots != 0)
                    {
                        output += sumSlots + " slots open in " + b.Name + "\n";
                    }
                }

                output += "A unscheduled: " + bestStudents.Where(x => x.Choices != null && x.A == null).Count() + "\n";
                output += "B unscheduled: " + bestStudents.Where(x => x.Choices != null && x.B == null).Count() + "\n";
                output += "Grade " + i + " score: " + score + " (lower is better)\n\n";
            }

            File.WriteAllLines("../../Sheets/Output/ConsoleOutput.txt", output.Split('\n'));

            Console.WriteLine(output);
            Console.ReadLine();
        }
    }
}
