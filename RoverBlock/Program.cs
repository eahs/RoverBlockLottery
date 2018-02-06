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
                // compute interest in specific rover blocks for grades 9 through 11
                Dictionary<String, int> choicesMap = new Dictionary<String, int>()
                {
                    { "Choice1", 4 },
                    { "Choice2", 5 },
                    { "Choice3", 6 },
                    { "Choice4", 7 },
                };
                rh.countChoices(choicesMap, i);



                // load students from master directed study list
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
                    { "Day", 0 },
                    { "BlockID", 1 },
                    { "NetworkID", 5 }
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
                dh.loadStudentChoices(studentChoiceMap, i, students);



                // output to XLS file
                sh.writeStudentsSheet(students, i);
            }



            /*
            
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
            
            */
        }
    }
}
