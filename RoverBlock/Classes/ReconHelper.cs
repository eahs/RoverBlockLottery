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

        // TODO: don't use Intersect when loading in choice data for interest, use distinct instead
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
                String NetworkID = entry["NetworkID"].ToLower().Replace("@roverkids.org", "");
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

        public void wallOfShame(Dictionary<String, int> map, int grade)
        {
            List<Student> students = new List<Student>();

            String fileName = "Choices" + grade + ".xls";
            List<Dictionary<String, String>> sheetData = sh.readSheet(fileName, map);

            foreach (Dictionary<String, String> entry in sheetData)
            {
                String NetworkID = entry["NetworkID"].ToLower().Replace("@roverkids.org", "");
                String LastName = entry["LastName"];
                String FirstName = entry["FirstName"];

                List<String> choices = new List<String>()
                {
                    entry["Choice1"],
                    entry["Choice2"],
                    entry["Choice3"],
                    entry["Choice4"]
                };

                if(choices.Distinct().Count() != 4)
                {
                    Student s = new Student(NetworkID, FirstName, LastName);
                    s.Choices = choices.GroupBy(x => x).Where(g => g.Count() > 1).Select(y => y.Key).ToList();
                    students.Add(s);
                }
            }

            sh.writeWallOfShame(students, grade);
        }
    }
}
