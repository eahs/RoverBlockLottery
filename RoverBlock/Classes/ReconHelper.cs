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

        public void countChoices(Dictionary<String, int> map, int grade)
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
                List<String> choices = new List<string>()
                {
                    entry["Choice1"],
                    entry["Choice2"],
                    entry["Choice3"],
                    entry["Choice4"]
                }.Distinct().ToList();

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
    }
}
