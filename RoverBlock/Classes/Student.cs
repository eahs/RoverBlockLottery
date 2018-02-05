using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoverBlock
{
    class Student
    {
        // id is the network id of the student. (eg. greenec)
        public String NetworkID;

        // the student's assigned classes
        public Block A;
        public Block B;

        // the student's class choices
        public List<String> Choices;

        public Student(String NetworkID, Block A, Block B)
        {
            this.NetworkID = NetworkID;
            this.A = A;
            this.B = B;
        }
    }
}
