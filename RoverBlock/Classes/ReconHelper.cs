using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoverBlock.Classes
{
    class ReconHelper
    {
        private static SheetHelper sh = new SheetHelper();

        public void countChoices(Dictionary<String, int> map, int grade, List<Student> students)
        {
            String fileName = "Choices" + grade + ".xls";
            List<Dictionary<String, int>> choiceCounts = new List<Dictionary<string, int>>()
            {
                new Dictionary<String, int>(),
                new Dictionary<String, int>(),
                new Dictionary<String, int>(),
                new Dictionary<String, int>()
            };

            List<Dictionary<String, String>> sheetData = sh.readSheet(fileName, map);
            foreach (Dictionary<String, String> entry in sheetData)
            {
                String NetworkID = entry["Email"].ToLower().Replace("@roverkids.org", "");
                Student s = students.Where(x => x.NetworkID == NetworkID).FirstOrDefault();

                if (s == null)
                {
                    continue;
                }

                // don't count a student's choices if they are locked into two classes already
                if(s.A != null && s.B != null)
                {
                    continue;
                }

                List<String> choices = s.Choices;

                for (int i = 0; i < choices.Count; i++)
                {
                    String className = choices[i];
                    if (!choiceCounts[i].ContainsKey(className))
                    {
                        choiceCounts[i].Add(className, 1);
                    }
                    else
                    {
                        choiceCounts[i][className]++;
                    }
                }
            }

            sh.writeChoicesSheet(choiceCounts, grade);
        }

        public void noChoiceStudents(List<Student> students, int grade)
        {
            List<Student> noChoices = new List<Student>();

            foreach (Student s in students)
            {
                // don't count a student's choices if they are locked into two classes already
                if (s.A != null && s.B != null)
                {
                    continue;
                }

                if(s.Choices == null)
                {
                    noChoices.Add(s);
                }
            }

            sh.writeNoChoicesSheet(noChoices, grade);
        }
    }
}
