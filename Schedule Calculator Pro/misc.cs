using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace Schedule_Calculator_Pro
{
    public class Don
    {
        public string donName { get; set; }
        public List<string> relatedSubjects { get; set; } = new List<string>();
        public string relatedAud { get; set; } = "";
        public bool[] possDays { get; set; } = new bool[] { true, true, true, true, true };
        public bool[][] dayStats { get; set; } = new bool[][] { new bool[] { true, true, true, true, true, true },
                                                                new bool[] { true, true, true, true, true, true },
                                                                new bool[] { true, true, true, true, true, true },
                                                                new bool[] { true, true, true, true, true, true },
                                                                new bool[] { true, true, true, true, true, true }};

        public Don(string donName)
        {
            this.donName = donName;
        }

        public void setExcludes(string vals)
        {
            vals = vals.Substring(1, vals.Length - 2);
            for (int x = 0; x < 5; x++)
                possDays[x] = vals[x] == '1';
            for (int x = 5; x < vals.Length; x++)
                dayStats[(x - 5) / 6][(x - 5) % 6] = vals[x] == '1';
        }

        public string getExcludes()
        {
            string r = "\"";
            for (int x = 0; x < 5; x++)
                r += (possDays[x]) ? 1 : 0;
            for (int x = 0; x < 5; x++)
                for (int y = 0; y < 6; y++)
                    r += (dayStats[x][y]) ? 1 : 0;
            r += '\"';
            return r;
        }

        public void excludeDay(string s)
        {
            if (s[s.Length - 1] == '/')
                s = s.Substring(0, s.Length - 1);
            int n = 0;
            switch (s)
            {
                case "Вівторок":
                    n = 1; break;
                case "Середа":
                    n = 2; break;
                case "Четвер":
                    n = 3; break;
                case "П'ятниця":
                    n = 4; break;
            }
            possDays[n] = false;
            dayStats[n] = new bool[] { false, false, false, false, false, false };
        }

        public void excludeDay(int n)
        {
            possDays[n] = false;
            dayStats[n] = new bool[] { false, false, false, false, false, false };
        }

        public void includeDay(string s)
        {
            int n = Program.DaynameToNum(s);
            possDays[n] = true;
            dayStats[n] = new bool[] { true, true, true, true, true, true };
        }

        public void fixConsD() // fix consistency w/ possDays as master
        {
            for(int x = 0; x < 5; x++)
            {
                if (possDays[x])
                {
                    dayStats[x] = new bool[] { true, true, true, true, true, true};
                }
                else
                {
                    dayStats[x] = new bool[] { false, false, false, false, false, false };
                }
            }
        }

        public void fixConsC() // fix consistency w/ dayStats as master
        {
            for (int x = 0; x < 5; x++)
            {
                possDays[x] = (isStrict(x) == 1 || isStrict(x) == 0);

            }
        }

        public int isStrict(int n)
        {
            var v = dayStats[n];
            return (v[0] && v[1] && v[2] && v[3] && v[4] && v[5]) ? 1 : (v[0] && !v[1] && !v[2] && !v[3] && !v[4] && !v[5]) ? -1 : 0; // if all v's are true 1, are false -1 and 0 if they're mixed
        }
    }

    public class Subject
    {
        public string subjectName { get; set; }
        public string relatedAud { get; set; } = "";
        public Subject(string subjectName)
        {
            this.subjectName = subjectName;
        }
    }

    public class Group
    {
        public string groupName { get; set; }
        public string relatedAud { get; set; } = "";
        public Dictionary<string, List<string>> relatedSubjects { get; set; } = new Dictionary<string, List<string>>();
        public int[] couplesXday { get; set; } = { -1, 3, -1, -1, 3 };
        public int StudyingWeeks { get; set; } = 16;
        public int course;
        public Group(string groupName)
        {
            this.groupName = groupName;
            couplesXday[0] = couplesXday[2] = couplesXday[3] = -1;
            couplesXday[1] = couplesXday[4] = 3;
            course = groupName[groupName.Length - 2] - '0';
        }
        public void SubjValEdit(string SubjName, string DonName, string SubjValue)
        {
            if (!relatedSubjects.ContainsKey(SubjName))
            {
                relatedSubjects.Add(SubjName, new List<string>() { DonName, SubjValue });
            }
            else
            {
                relatedSubjects[SubjName].Add(DonName);
                relatedSubjects[SubjName][1] = SubjValue;
            }
        }

        public int couplesXdayGet(int day)      // Возвращает кол-во пар для конкретного дня и вызывает их рассчёт, если словарь пуст.
        {
            if (couplesXday.Contains(-1))
                couplesXdayCalc();
            return couplesXday[day];
        }


        public void couplesXdayCalc()            // Рассчёт кол-ва пар на каждый день
        {
            // Важная формула: (relatedSubjects.Sum(x => x.Value) - 3 * StudyingWeeks) / (StudyingWeeks * 4);
            int uncalcdays = 4, modifier = 3 * StudyingWeeks;
            var t1 = couplesXday;
            while (uncalcdays > 0)
            {
                double tmp = (relatedSubjects.Sum(x => Convert.ToInt32(x.Value[1])) - Convert.ToDouble(modifier)) / (Convert.ToDouble(StudyingWeeks) * Convert.ToDouble(uncalcdays));
                if (tmp < Math.Ceiling(tmp))
                {
                    if (tmp <= 3)
                    {
                        for (int x = 0; x < 4; x++)
                            if (t1[x] == -1)
                                couplesXday[x] = 3;
                        break;
                    }
                    if (uncalcdays == 1)
                    {
                        if (t1[0] >= 4 && t1[2] >= 4 && t1[3] >= 4)
                            couplesXday[1] = Convert.ToInt32(Math.Ceiling(tmp));
                        break;
                    }
                    modifier += Convert.ToInt32(Math.Ceiling(tmp)) * StudyingWeeks;
                    couplesXday[Array.IndexOf(couplesXday, -1)] = Convert.ToInt32(Math.Ceiling(tmp));
                    uncalcdays--;
                }
                else
                {
                    for (int x = 0; x < 4; x++)
                        if (couplesXday[x] == -1 || x == 1)
                            couplesXday[x] = Convert.ToInt16(tmp);
                    break;
                }
            }
        }
    }
}
