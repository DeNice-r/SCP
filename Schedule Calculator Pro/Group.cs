using System;
using System.Collections.Generic;
using System.Linq;

namespace Schedule_Calculator_Pro
{
    public class Group
    {
        public string groupName { get; set; }
        public string relatedAud { get; set; } = "";
        public Dictionary<string, List<string>> relatedSubjects { get; set; } = new Dictionary<string, List<string>>();
        public Dictionary<string, int> relatedSubjectsx2 { get; set; } = new Dictionary<string, int>();
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
            if (!relatedSubjectsx2.ContainsKey(SubjName))
                relatedSubjectsx2.Add(SubjName, 0);
        }

        public int couplesXdayGet(int day)      
        {
            if (couplesXday.Contains(-1))
                couplesXdayCalc();
            return couplesXday[day];
        }

        public void couplesXdayCalc()  
        {
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