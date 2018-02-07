using RoverBlock.Classes;
using System;
using System.Collections.Generic;
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



            for (int i = 9; i < 12; i++)
            {
                int score = int.MaxValue;
                List<Student> bestStudents = new List<Student>();
                List<Block> bestBlocks = new List<Block>();

                for (int j = 0; j < 1000; j++) {

                    Dictionary<String, int> blocksMap = new Dictionary<string, int>()
                    {
                        { "Block ID", 0 },
                        { "Block Name", 1 },
                        { "A Slots", 2 },
                        { "B Slots", 3 }
                    };
                    List<Block> blocks = dh.getBlocks(blocksMap, i);



                    // load students from master directed study lists
                    Dictionary<String, int> studentsMap = new Dictionary<string, int>()
                    {
                        { "NetworkID", 0 },
                        { "LastName", 1 },
                        { "FirstName", 2 }
                    };
                    List<Student> students = dh.getStudents(studentsMap, i);



                    // lock students into specific rover blocks
                    Dictionary<String, int> lockedStudentsMap = new Dictionary<String, int>()
                    {
                        { "Day", 4 },
                        { "BlockID", 1 },
                        { "NetworkID", 3 }
                    };
                    dh.lockStudents("LockedStudents.xls", lockedStudentsMap, students);



                    // load the students' choices into a list and remove duplicates
                    Dictionary<String, int> studentChoiceMap = new Dictionary<String, int>()
                    {
                        { "NetworkID", 8 },
                        { "Choice1", 4 },
                        { "Choice2", 5 },
                        { "Choice3", 6 },
                        { "Choice4", 7 }
                    };
                    dh.loadStudentChoices(studentChoiceMap, i, students, blocks);



                    // compute interest in specific rover blocks for grades 9 through 11
                    Dictionary<String, int> choicesMap = new Dictionary<String, int>()
                    {
                        { "Email", 8},
                        { "Choice1", 4 },
                        { "Choice2", 5 },
                        { "Choice3", 6 },
                        { "Choice4", 7 },
                    };
                    rh.countChoices(choicesMap, i, students);



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



                    int newScore = students.Where(x => x.Choices != null && x.A == null).Count() + students.Where(x => x.B == null).Count();

                    if(newScore < score)
                    {
                        bestStudents = students.ToList();
                        bestBlocks = blocks.ToList();
                        score = newScore;
                    }
                }

                // output to XLS file
                sh.writeStudentsSheet(bestStudents, i);

                foreach (Block b in bestBlocks)
                {
                    int sumSlots = b.aSlots + b.bSlots;

                    if(sumSlots != 0)
                    {
                        Console.WriteLine(sumSlots + " slots open in " + b.Name);
                    }
                }

                Console.WriteLine("Grade " + i + " score: " + score + " (lower is better)\n");
            }
            
            Console.ReadLine();
        }
    }
}
