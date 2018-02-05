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

            Dictionary<String, int> lockedStudentsMap = new Dictionary<string, int>()
            {
                { "Day", 0 },
                { "NetworkID", 1 }
            };
            List<Student> lockedStudents = dh.getLockedStudents("LockedStudents.xls", lockedStudentsMap);



            Dictionary<String, int> choicesMap = new Dictionary<string, int>()
            {
                { "NetworkID", 0 },
                { "Choice 1", 1 },
                { "Choice 2", 2 },
                { "Choice 3", 3 }
            };
            List<Student> students = dh.getStudents("Choices.xls", choicesMap, lockedStudents);

            // free up some memory here
            lockedStudents = null;
            GC.Collect();



            Dictionary<String, int> blocksMap = new Dictionary<string, int>()
            {
                { "Class Name", 0 },
                { "A Slots", 1 },
                { "B Slots", 2 }
            };
            List<Block> blocks = dh.getBlocks("Classes.xls", blocksMap);

            foreach(Block b in blocks)
            {
                dh.runLotteryA(b, students);
                dh.runLotteryB(b, students);
            }

            Console.WriteLine(String.Join(", ", blocks.Select(x => x.Name)) + "\n");



            foreach (Student s in students)
            {
                Console.WriteLine("ID: " + s.NetworkID);
                Console.WriteLine("A: " + (s.A == null ? "null" : s.A.Name));
                Console.WriteLine("B: " + (s.B == null ? "null" : s.B.Name));
                Console.WriteLine(String.Join(", ", s.Choices));
                Console.WriteLine();
            }



            dh.writeStudentsSheet(students);



            Console.ReadLine();
        }
    }
}
