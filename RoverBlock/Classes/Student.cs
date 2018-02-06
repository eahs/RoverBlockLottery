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
        public String FirstName;
        public String LastName;

        // the student's assigned classes
        public Block A;
        public Block B;

        // the student's class choices
        public List<String> Choices;

        public Student(String NetworkID, String FirstName, String LastName)
        {
            this.NetworkID = NetworkID;
            this.FirstName = FirstName;
            this.LastName = LastName;
        }
    }
}
