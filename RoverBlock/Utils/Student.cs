using System;
using System.Collections.Generic;

namespace RoverBlock
{
    [Serializable]
    public class Student
    {
        public string NetworkID { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }

        public int LotteryScore { get; set; } = 0;

        // the student's assigned classes
        public Block RoverBlock { get; set; }

        // the student's class choices
        public List<string> Choices;
    }
}
