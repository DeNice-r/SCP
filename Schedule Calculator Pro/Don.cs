using System.Collections.Generic;
using System.Linq;

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
            if (vals == "" || vals == "\"\"")
                vals = "\"11111111111111111111111111111111111\"";
            if (vals.Contains('"'))
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

        public void fixConsD()
        {
            for (int x = 0; x < 5; x++)
            {
                if (possDays[x])
                {
                    dayStats[x] = new bool[] { true, true, true, true, true, true };
                }
                else
                {
                    dayStats[x] = new bool[] { false, false, false, false, false, false };
                }
            }
        }

        public void fixConsC()
        {
            for (int x = 0; x < 5; x++)
            {
                possDays[x] = (isStrict(x) == 1 || isStrict(x) == 0);
            }
        }

        public int isStrict(int n)
        {
            var v = dayStats[n];
            return (v[0] && v[1] && v[2] && v[3] && v[4] && v[5]) ? 1 : (!v[0] && !v[1] && !v[2] && !v[3] && !v[4] && !v[5]) ? -1 : 0;
        }
    }
}