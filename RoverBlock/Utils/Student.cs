using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoverBlock
{
    [Serializable]
    public class Student
    {
        // id is the network id of the student. (eg. greenec)
        public string NetworkID;
        public string FirstName;
        public string LastName;

        // the student's assigned classes
        public Block A;
        public Block B;

        // the student's class choices
        public List<string> Choices;

        public Student(string NetworkID, string FirstName, string LastName)
        {
            this.NetworkID = NetworkID;
            this.FirstName = FirstName;
            this.LastName = LastName;
        }
    }
}
