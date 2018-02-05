﻿using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RoverBlock.Classes
{
    class DataHelper
    {
        private static Random rnd = new Random();

        public List<Student> getLockedStudents(String fileName, Dictionary<String, int> map)
        {
            List<LockedStudent> lockedStudents = new List<LockedStudent>();

            HSSFWorkbook hssfwb;
            using (FileStream file = new FileStream("../../Sheets/" + fileName, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheetAt(0);
            for (int row = 1; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null)
                {
                    List<ICell> cells = sheet.GetRow(row).Cells;

                    String NetworkID = cells[map["NetworkID"]].ToString();
                    String Day = cells[map["Day"]].ToString();

                    lockedStudents.Add(new LockedStudent(NetworkID, Day));
                }
            }

            Block reservedBlock = new Block("RESERVED", 0, 0);

            List<Student> output = (
                from ls in lockedStudents
                join lsA in lockedStudents
                    on new { a = ls.NetworkID, b = "A" } equals new { a = lsA.NetworkID, b = lsA.Day } into joinedA
                from lsA in joinedA.DefaultIfEmpty()
                join lsB in lockedStudents
                    on new { a = ls.NetworkID, b = "B" } equals new { a = lsB.NetworkID, b = lsB.Day } into joinedB
                from lsB in joinedB.DefaultIfEmpty()
                select new Student(ls.NetworkID, lsB == null ? null : reservedBlock, lsA == null ? null : reservedBlock)
            ).GroupBy(x => x.NetworkID).Select(x => x.FirstOrDefault()).ToList();

            return output;
        }

        public List<Student> getStudents(String fileName, Dictionary<String, int> map, List<Student> lockedStudents)
        {
            List<Student> students = new List<Student>();

            HSSFWorkbook hssfwb;
            using (FileStream file = new FileStream("../../Sheets/" + fileName, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheetAt(0);
            for (int row = 1; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null)
                {
                    List<ICell> cells = sheet.GetRow(row).Cells;

                    String NetworkID = cells[map["NetworkID"]].ToString();
                    List<String> choices = new List<String>()
                    {
                        cells[map["Choice 1"]].ToString(),
                        cells[map["Choice 2"]].ToString(),
                        cells[map["Choice 3"]].ToString()
                    }.Distinct().ToList();

                    // TODO: use an intersect here. that would solve the duplicate issue and "promote" choices if the class does not exist

                    // reuse Student object from locked students to preserve RESERVED classes
                    Student s = lockedStudents.Where(x => x.NetworkID == NetworkID).SingleOrDefault();

                    // if no student matched, create a new one
                    if (s == null)
                    {
                        s = new Student(NetworkID, null, null);
                    }

                    s.Choices = choices;

                    students.Add(s);
                }
            }

            return students;
        }

        public List<Block> getBlocks(String fileName, Dictionary<String, int> map)
        {
            List<Block> blocks = new List<Block>();

            HSSFWorkbook hssfwb;
            using (FileStream file = new FileStream("../../Sheets/" + fileName, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheetAt(0);
            for (int row = 1; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null)
                {
                    List<ICell> cells = sheet.GetRow(row).Cells;

                    String BlockName = cells[map["Class Name"]].ToString();
                    int aSlots = tryParse(cells[map["A Slots"]].ToString());
                    int bSlots = tryParse(cells[map["B Slots"]].ToString());

                    blocks.Add(new Block(BlockName, aSlots, bSlots));
                }
            }

            shuffle(blocks);
            return blocks;
        }

        public void runLotteryA(Block b, List<Student> students)
        {
            List<Student> interestedStudents = students.Where(x => x.A == null && x.Choices.Contains(b.Name)).ToList();
            List<Student> pool = new List<Student>();

            foreach (Student s in interestedStudents)
            {
                pool.AddRange(Enumerable.Repeat(s, 4 - s.Choices.IndexOf(b.Name)));
            }

            interestedStudents = null;
            GC.Collect();

            // TODO: check behavior for underfilled classes
            for (int i = 0; i < b.aSlots; i++)
            {
                if(pool.Count == 0)
                {
                    break;
                }

                Student winner = pool[rnd.Next(pool.Count)];
                winner.Choices.Remove(b.Name);
                winner.A = b;
                pool.RemoveAll(x => x.NetworkID == winner.NetworkID);
            }
        }

        public void runLotteryB(Block b, List<Student> students)
        {
            List<Student> interestedStudents = students.Where(x => x.B == null && x.Choices.Contains(b.Name)).ToList();
            List<Student> pool = new List<Student>();

            foreach (Student s in interestedStudents)
            {
                pool.AddRange(Enumerable.Repeat(s, 4 - s.Choices.IndexOf(b.Name)));
            }

            interestedStudents = null;
            GC.Collect();

            // TODO: check behavior for underfilled classes
            for (int i = 0; i < b.bSlots; i++)
            {
                if (pool.Count == 0)
                {
                    break;
                }

                Student winner = pool[rnd.Next(pool.Count)];
                winner.Choices.Remove(b.Name);
                winner.B = b;
                pool.RemoveAll(x => x.NetworkID == winner.NetworkID);
            }
        }

        public int tryParse(String str)
        {
            return int.TryParse(str, out int result) ? result : 0;
        }

        public static void shuffle<T>(List<T> list)
        {
            int n = list.Count;
            while (n > 1)
            {
                n--;
                int k = rnd.Next(n + 1);
                T value = list[k];
                list[k] = list[n];
                list[n] = value;
            }
        }
    }
}
