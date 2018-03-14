using System.Collections.Generic;
using System.Linq;

namespace RoverBlock
{
    public static class ReconHelper
    {
        // TODO: don't use Intersect when loading in choice data for interest, use distinct instead
        public static void CountChoices(Dictionary<string, int> map, int grade, List<Student> students)
        {
            string fileName = "Choices" + grade + ".xls";
            List<Dictionary<string, int>> choiceCounts = new List<Dictionary<string, int>>()
            {
                new Dictionary<string, int>(),
                new Dictionary<string, int>(),
                new Dictionary<string, int>(),
                new Dictionary<string, int>()
            };

            List<Dictionary<string, string>> sheetData = SheetHelper.ReadSheet(fileName, map);
            foreach (Dictionary<string, string> entry in sheetData)
            {
                string NetworkID = entry["NetworkID"].ToLower().Replace("@roverkids.org", "");
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

                List<string> choices = s.Choices;

                for (int i = 0; i < choices.Count; i++)
                {
                    string className = choices[i];
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

            SheetHelper.WriteChoicesSheet(choiceCounts, grade);
        }

        public static void NoChoiceStudents(List<Student> students, int grade)
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

            SheetHelper.WriteNoChoicesSheet(noChoices, grade);
        }

        public static void WallOfShame(Dictionary<string, int> map, int grade)
        {
            List<Student> students = new List<Student>();

            string fileName = "Choices" + grade + ".xls";
            List<Dictionary<string, string>> sheetData = SheetHelper.ReadSheet(fileName, map);

            foreach (Dictionary<string, string> entry in sheetData)
            {
                string NetworkID = entry["NetworkID"].ToLower().Replace("@roverkids.org", "");
                string LastName = entry["LastName"];
                string FirstName = entry["FirstName"];

                List<string> choices = new List<string>()
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

            SheetHelper.WriteWallOfShame(students, grade);
        }
    }
}
