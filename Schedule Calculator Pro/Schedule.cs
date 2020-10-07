using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Text.Json;
using System.Text.Json.Serialization;


namespace Schedule_Calculator_Pro
{
    public class Schedule
    {
        public List<List<List<List<string>>>> schedule { get; set; } = new List<List<List<List<string>>>>();        // Группа > День > Пара > Преподаватели/Предметы/Аудитории
        public List<List<List<List<string>>>> scheduleFree { get; set; } = new List<List<List<List<string>>>>();        // День > Пара >  Преподаватель/Аудитория
        private Random rnd = new Random();
        public Schedule() { }

        public void Start()
        {
            PreGen();
            var pdon = Program.don;
            var pgroup = Program.group;
            for (int don = 0; don < pdon.Count; don++)
            {
                var tgroups = new List<int>();
                var tlist = new List<List<int>>();
                for (int group = 0; group < pgroup.Count; group++)
                {
                    var tgroup = pgroup.Values.ToArray()[group];
                    for (int x = 0; x < tgroup.relatedSubjects.Count; x++)
                    {
                        if (tgroup.relatedSubjects.Values.ToArray()[x][0] == pdon.Values.ToArray()[don].donName)
                        {
                            if (!tgroups.Contains(group))
                                tgroups.Add(group);
                            if (tlist.Count < tgroups.Count)
                                tlist.Add(new List<int>());
                            tlist[tlist.Count - 1].Add(x);
                        }
                    }
                }
                for (int day = 0; day < 5; day++)
                {

                    for (int couple = 0; couple < 6; couple++)
                    {
                        // Если преподаватель свободен в этот день на этой паре
                        if (scheduleFree[day][couple][0].Contains(pdon.Keys.ToArray()[don]))
                            for (int group = 0; group < tgroups.Count; group++)
                            {
                                if (pgroup.Values.ToArray()[tgroups[group]].couplesXdayGet(day) <= couple || schedule[tgroups[group]][day][couple].Count != 0)
                                    continue;
                                for (int subj = 0; subj < tlist[group].Count; subj++)
                                    if (Convert.ToInt32(pgroup.Values.ToArray()[tgroups[group]].relatedSubjects.Values.ToArray()[tlist[group][subj]][1]) > 0)
                                    {
                                        schedule[tgroups[group]][day][couple].Add(pgroup.Values.ToArray()[tgroups[group]].relatedSubjects.Keys.ToArray()[tlist[group][subj]]); // Предмет
                                        schedule[tgroups[group]][day][couple].Add(pdon.Keys.ToArray()[don]); // Преподаватель
                                        var b = pgroup.Values.ToArray()[tgroups[group]].relatedSubjects.Values.ToArray()[tlist[group][subj]].Count == 3 && pgroup.Values.ToArray()[tgroups[group]].relatedSubjects.Values.ToArray()[tlist[group][subj]][2] != "";
                                        if (b)
                                        {
                                            schedule[tgroups[group]][day][couple].Add(pgroup.Values.ToArray()[tgroups[group]].relatedSubjects.Values.ToArray()[tlist[group][subj]][2]);
                                            scheduleFree[day][couple][0].Remove(pgroup.Values.ToArray()[tgroups[group]].relatedSubjects.Values.ToArray()[tlist[group][subj]][2]);
                                        }
                                        schedule[tgroups[group]][day][couple].Add(getAud(day, couple, pgroup.Values.ToArray()[tgroups[group]], pgroup.Values.ToArray()[tgroups[group]].relatedSubjects.Keys.ToArray()[tlist[group][subj]], 0)); // Аудитория
                                        if (b)
                                            schedule[tgroups[group]][day][couple].Add(getAud(day, couple, pgroup.Values.ToArray()[tgroups[group]], pgroup.Values.ToArray()[tgroups[group]].relatedSubjects.Keys.ToArray()[tlist[group][subj]], 0)); // Аудитория
                                        scheduleFree[day][couple][0].Remove(pdon.Keys.ToArray()[don]);
                                        pgroup.Values.ToArray()[tgroups[group]].relatedSubjects.Values.ToArray()[tlist[group][subj]][1] = (Convert.ToInt32(pgroup.Values.ToArray()[tgroups[group]].relatedSubjects.Values.ToArray()[tlist[group][subj]][1]) - 16).ToString();
                                        group = 10000;
                                        break;
                                    }
                            }
                    }
                }
            }
            Save();
            MessageBox.Show("Розклад складено.");
        }

        private void halfObjRem(int day, int couple, int idx, string obj)
        {
            if (obj.Contains("/"))
            {
                scheduleFree[day][couple][idx].Remove(obj);
            }
            else
            {
                scheduleFree[day][couple][idx].Remove(obj);
                scheduleFree[day][couple][idx].Add(obj + "/");
            }
        }

        private string getAud(int day, int couple, Group group, string subject, int way)
        {
            var t = "";
            if (Program.subject[subject].relatedAud != "" && (scheduleFree[day][couple][1].Contains(Program.subject[subject].relatedAud) || scheduleFree[day][couple][1].Contains(Program.subject[subject].relatedAud + "/") || Program.subject[subject].relatedAud.Contains("с.з.")))
                t = Program.subject[subject].relatedAud;
            else if (group.relatedAud != "" && scheduleFree[day][couple][1].Contains(group.relatedAud))
                t = group.relatedAud;
            else if (scheduleFree[day][couple][1].Contains(Program.don[group.relatedSubjects[subject][0]].relatedAud))
                t = Program.don[group.relatedSubjects[subject][0]].relatedAud;
            else
            {
                if (scheduleFree[day][couple][1].Count < 2)
                {
                    MessageBox.Show("Недостатньо аудиторій для розміщення всіх груп!");
                    Program.SchedGenThread.Abort();
                }
                t = scheduleFree[day][couple][1][rnd.Next(0, scheduleFree[day][couple][1].Count)];
            }
            if (way == 0)
                scheduleFree[day][couple][1].Remove(t);
            else
                halfObjRem(day, couple, 1, t);
            return t;
        }

        private void PreGen()
        {
            scheduleFree = new List<List<List<List<string>>>>();
            schedule = new List<List<List<List<string>>>>();
            var donk = Program.don.Keys.ToList();
            var donv = Program.don.Values.ToList();
            for (int _day = 0; _day < 5; _day++)
            {
                scheduleFree.Add(new List<List<List<string>>>());
                for (int _couple = 0; _couple < 6; _couple++)
                {

                    scheduleFree[_day].Add(new List<List<string>>());
                    scheduleFree[_day][_couple].Add(new List<string>());
                    for (int don = 0; don < donk.Count; don++)
                    {
                        if (donv[don].possDays[_day] && donv[don].dayStats[_day][_couple])
                        {
                            scheduleFree[_day][_couple][0].Add(donk[don]);
                        }
                    }
                    scheduleFree[_day][_couple].Add(new List<string>());
                    scheduleFree[_day][_couple][1] = Program.audience.ToList();
                }
            }
            for (int x = 0; x < Program.group.Count; x++)
            {
                schedule.Add(new List<List<List<string>>>());
                if (Program.group.Values.ToArray()[x].couplesXday.Contains(-1))
                    Program.group.Values.ToArray()[x].couplesXdayCalc();
                for (int y = 0; y < 5; y++)
                {
                    schedule[x].Add(new List<List<string>>());
                    for (int z = 0; z < Program.group.Values.ToArray()[x].couplesXday[y]; z++)
                        schedule[x][y].Add(new List<string>());
                }
            }
        }

        public void Save()
        {
            // Save in json to work w/ it from program later.
            
            System.IO.File.WriteAllText(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\schedcfg.json"), JsonSerializer.Serialize(schedule));
            System.IO.File.WriteAllText(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\schedfreecfg.json"), JsonSerializer.Serialize(scheduleFree));

            for (int _course = 0; _course < 4; _course++)
            {

                var xcl = new Excel(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\Розклад " + (_course + 1) + " курс.xlsx"));
                var col = 2;
                var row = 1;
                var mxcpl = new List<int>() { 0, 0, 0, 0, 0 };

                for (int dy = 0; dy < 5; dy++)
                {
                    for (int grp = 0; grp < schedule.Count; grp++)
                    {
                        var t = 0;
                        if (Program.group.Values.ToArray()[grp].course == _course + 1 && mxcpl[dy] < schedule[grp][dy].Count)
                            for (int z = 0; z < schedule[grp][dy].Count; z++)
                                if (schedule[grp][dy][z].Count == 5)
                                    t++;
                        mxcpl[dy] = schedule[grp][dy].Count + t;
                    }
                }

                var wtf = new List<List<List<int>>>();

                for (int crs = 0; crs < 4; crs++)
                {
                    wtf.Add(new List<List<int>>());
                    for (int dy = 0; dy < 5; dy++)
                    {
                        wtf[crs].Add(new List<int>());
                        for (int cpl = 0; cpl < 6; cpl++)
                        {
                            for (int grp = 0; grp < Program.group.Count; grp++)
                            {
                                if (Program.group.Values.ToArray()[grp].course - 1 != crs)
                                    continue;
                                if (schedule[grp][dy].Count < cpl)
                                    wtf[crs][dy].Add(0);
                            }
                        }
                    }
                }

                for (int crs = 0; crs < 4; crs++)
                {
                    for (int dy = 0; dy < 5; dy++)
                    {
                        for (int cpl = 0; cpl < 6; cpl++)
                        {
                            for (int grp = 0; grp < Program.group.Count; grp++)
                            {
                                if (Program.group.Values.ToArray()[grp].course - 1 != crs)
                                    continue;
                                if (wtf[crs][dy][cpl] == 2)
                                    continue;
                                if (schedule[grp][dy].Count <= cpl)
                                    continue;
                                if (schedule[grp][dy][cpl].Count == 5)
                                    wtf[crs][dy][cpl] = 2;
                                else if (schedule[grp][dy][cpl].Count == 3)
                                    wtf[crs][dy][cpl] = 1;
                            }
                        }
                    }
                }

                for (int _group = 0; _group < schedule.Count; _group++)
                {
                    var mx = -1;
                    if (Program.group.Values.ToArray()[_group].course == _course + 1)
                    {
                        xcl.ws.Range[getcolname(row + 1) + "2", getcolname(row + 3) + "2"].Merge();
                        xcl.WriteToCell(1, row, Program.group.Keys.ToArray()[_group]);
                        xcl.ws.Range[getcolname(1 + 1) + (row + 1).ToString()].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        for (int _day = 0; _day < schedule[_group].Count; _day++)
                        {
                            var tcol = col;
                            for (int _couple = 0; _couple < schedule[_group][_day].Count; _couple++)
                            {
                                var trow = row;
                                if (schedule[_group][_day][_couple].Count == 3)
                                {
                                    //for (int g = 0; g < Program.group.Count; g++)
                                    //{
                                    //    var t = 0;
                                    //    if (_day != 0)
                                    //    {
                                    //        t = func1(wtf[_course][_day - 1], _course, _day);
                                    //        while (t + 1 > schedule[g][_day - 1].Count) { t--; }
                                    //    }
                                    //    if (xcl.BReadCell(tcol, trow))
                                    //    {
                                    //        tcol++;
                                    //        break;
                                    //    }
                                    //}
                                    for (int c = 0; c < schedule[_group][_day][_couple].Count; c++)
                                    {
                                        xcl.WriteToCell(tcol, trow, schedule[_group][_day][_couple][c]);
                                        //if (mrow < trow)
                                        //    mrow = trow;
                                        //if (mcol < tcol)
                                        //    mcol = tcol;
                                        trow++;
                                    }
                                }
                                else if (schedule[_group][_day][_couple].Count == 5)
                                {
                                    //for (int g = 0; g < Program.group.Count; g++)
                                    //{
                                    //    var t = 0;
                                    //    if (_day != 0)
                                    //    {
                                    //        t = func1(wtf[_course][_day - 1], _course, _day);
                                    //        while (t + 1 > schedule[g][_day - 1].Count) { t--; }
                                    //    }
                                    //    if (Program.group.Values.ToArray()[g].course == _course + 1 && ((_couple != 0 && schedule[g][_day][_couple - 1].Count == 5) || (_day != 0 && schedule[g][_day - 1][t].Count == 5)))
                                    //    {
                                    //        tcol++;
                                    //        break;
                                    //    }
                                    //}
                                    xcl.ws.Range[getcolname(trow + 1) + (tcol + 1).ToString() + ":" + getcolname(trow + 1) + (tcol + 2).ToString()].Merge();
                                    xcl.ws.get_Range(getcolname(trow + 1) + (tcol + 1).ToString() + ":" + getcolname(trow + 1) + (tcol + 2).ToString()).Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    xcl.WriteToCell(tcol, trow, schedule[_group][_day][_couple][0]); trow++;
                                    xcl.WriteToCell(tcol, trow, schedule[_group][_day][_couple][1]); trow++;
                                    xcl.WriteToCell(tcol, trow, schedule[_group][_day][_couple][3]); trow = row + 1; tcol++; mxcpl[_day]++;
                                    xcl.WriteToCell(tcol, trow, schedule[_group][_day][_couple][2]); trow++;
                                    xcl.WriteToCell(tcol, trow, schedule[_group][_day][_couple][4]);
                                    tcol--; //
                                    //if (mrow < trow)
                                    //    mrow = trow;
                                    //if (mcol < tcol)
                                    //    mcol = tcol;
                                    trow++;
                                }
                                if (mx < trow - row)
                                    mx = trow - row;
                                tcol += wtf[_course][_day][_couple];
                            }
                            //col += mxcpl[_day]+1;
                            col += wtf[_course][_day].Sum();
                        }
                        if (_group != schedule.Count - 1)
                            col = 2;
                        row += mx;
                    }
                    if (_group != schedule.Count - 1)
                        col = 2;
                }
                xcl.ws.Range[xcl.ws.Cells[1, 1], xcl.ws.Cells[1, Program.group.Count(x => x.Value.course == _course + 1) * 3 + 1]].Merge();
                var s1 = "";
                switch (_course) { case 0: { s1 = "I"; } break; case 1: { s1 = "II"; } break; case 2: { s1 = "III"; } break; case 3: { s1 = "IV"; } break; }
                s1 += " курс";
                xcl.WriteToCell(0, 0, s1);
                xcl.ws.get_Range("A1").Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                var t3 = Program.group.Count(x => x.Value.course - 1 == _course) * 3;
                var t2 = wtf[_course].Sum(x => x.Sum()) + 1;
                xcl.ws.Range["A1:" + getcolname(t3 + 1) + (1 + t2).ToString()].Borders.LineStyle = XlLineStyle.xlContinuous;
                xcl.ws.Range["A1:" + getcolname(t3 + 1) + (1 + t2).ToString()].Borders.Weight = XlBorderWeight.xlThin;
                xcl.ws.Range["A1:" + getcolname(t3 + 1) + (1 + t2).ToString()].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                xcl.ws.Range["A1:" + getcolname(t3 + 1) + (1 + t2).ToString()].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
                xcl.ws.Range["A1:" + getcolname(t3 + 1) + (1 + t2).ToString()].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                xcl.ws.Range["A1:" + getcolname(t3 + 1) + (1 + t2).ToString()].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium; // mcol => wtf[_course].Sum(x => x.Sum())
                int t1 = 3;
                for (int zz = 0; zz < 5; zz++)
                {
                    xcl.ws.Range[xcl.ws.Cells[t1, 1], xcl.ws.Cells[t1 + wtf[_course][zz].Sum() - 1, 1]].Merge();
                    switch (zz)
                    {
                        case 0:
                            {
                                xcl.WriteToCell(t1 - 1, 0, "Понеділок");

                            }
                            break;
                        case 1:
                            {
                                xcl.WriteToCell(t1 - 1, 0, "Вівторок");
                            }
                            break;
                        case 2:
                            {
                                xcl.WriteToCell(t1 - 1, 0, "Середа");
                            }
                            break;
                        case 3:
                            {
                                xcl.WriteToCell(t1 - 1, 0, "Четвер");
                            }
                            break;
                        case 4:
                            {
                                xcl.WriteToCell(t1 - 1, 0, "П'ятниця");
                            }
                            break;
                    }

                    xcl.ws.Range[xcl.ws.Cells[t1, 1], xcl.ws.Cells[t1 + wtf[_course][zz].Sum() - 1, (Program.group.Count(x => x.Value.course == _course + 1) * 3 + 1)]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                    xcl.ws.Range[xcl.ws.Cells[t1, 1], xcl.ws.Cells[t1 + wtf[_course][zz].Sum() - 1, (Program.group.Count(x => x.Value.course == _course + 1) * 3 + 1)]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
                    xcl.ws.Range[xcl.ws.Cells[t1, 1], xcl.ws.Cells[t1 + wtf[_course][zz].Sum() - 1, (Program.group.Count(x => x.Value.course == _course + 1) * 3 + 1)]].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                    xcl.ws.Range[xcl.ws.Cells[t1, 1], xcl.ws.Cells[t1 + wtf[_course][zz].Sum() - 1, (Program.group.Count(x => x.Value.course == _course + 1) * 3 + 1)]].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
                    xcl.ws.Range[xcl.ws.Cells[t1, 1], xcl.ws.Cells[t1 + wtf[_course][zz].Sum() - 1, 1]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    xcl.ws.Range[xcl.ws.Cells[t1, 1], xcl.ws.Cells[t1 + wtf[_course][zz].Sum() - 1, 1]].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    t1 += wtf[_course][zz].Sum();
                }
                for (int i = 2; i < 28; i++)
                {
                    xcl.ws.Cells[i, 1].Font.Bold = true;
                }

                xcl.ws.get_Range("B2", "BP50").Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                xcl.ws.Cells[1, 1].Font.Bold = true;
                xcl.ws.Cells[1, 1].Font.Size = 20;
                xcl.ws.Columns.AutoFit();
                xcl.SaveAs();
                xcl.close();
            }
            Excel xcl1 = new Excel(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\Вільні викладачі.xlsx"));
            var xrow = 1;
            var xcol = 0;
            var xrowt = 0;

            UInt16[] cxz = new UInt16[5] { 0, 0, 0, 0, 0 };

            for (int y = 0; y < 5; y++)
            {
                for (int x = 0; x < 6; x++)
                {
                    if (scheduleFree[y][x][0].Count > cxz[y])
                        cxz[y] = Convert.ToUInt16(scheduleFree[y][x][0].Count);
                }
            }
            for (int x = 0; x < 5; x++)
            {
                switch (x) { case 0: { xcl1.WriteToCell(0, 0, "Понеділок"); } break; case 1: { xcl1.WriteToCell(xrow, xcol, "Вівторок"); } break; case 2: { xcl1.WriteToCell(xrow, xcol, "Середа"); } break; case 3: { xcl1.WriteToCell(xrow, xcol, "Четвер"); } break; case 4: { xcl1.WriteToCell(xrow, xcol, "П'ятниця"); } break; }
                if (x == 0)
                {
                    xcl1.ws.Range["A" + (xrow).ToString(), "F" + (xrow).ToString()].Merge();
                    xcl1.ws.Range["A" + (xrow).ToString(), "F" + (xrow).ToString()].Font.Bold = true;
                    xcl1.ws.Range["A" + (xrow).ToString(), "F" + (xrow).ToString()].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range["A" + (xrow).ToString(), "F" + (xrow).ToString()].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range["A" + (xrow).ToString(), "F" + (xrow).ToString()].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range["A" + (xrow).ToString(), "F" + (xrow).ToString()].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range["A" + (xrow).ToString(), "F" + (xrow).ToString()].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                }
                else
                {
                    xcl1.ws.Range["A" + (xrow + 1).ToString(), "F" + (xrow + 1).ToString()].Merge();
                    xcl1.ws.Range["A" + (xrow + 1).ToString(), "F" + (xrow + 1).ToString()].Font.Bold = true;
                    xcl1.ws.Range["A" + (xrow + 1).ToString(), "F" + (xrow + 1).ToString()].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range["A" + (xrow + 1).ToString(), "F" + (xrow + 1).ToString()].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range["A" + (xrow + 1).ToString(), "F" + (xrow + 1).ToString()].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range["A" + (xrow + 1).ToString(), "F" + (xrow + 1).ToString()].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range["A" + (xrow + 1).ToString(), "F" + (xrow + 1).ToString()].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                }
                xrow = 2;
                for (int tx = 0; tx < x; tx++)
                {
                    xrow += cxz[tx] + 2;
                }
                for (int y = 0; y < 6; y++)
                {
                    xrow = 1;
                    for (int tx = 0; tx < x; tx++)
                    {
                        xrow += cxz[tx] + 2;
                    }
                    var s = "";
                    switch (y) { case 0: { s = "Перша"; } break; case 1: { s = "Друга"; } break; case 2: { s = "Третя"; } break; case 3: { s = "Четверта"; } break; case 4: { s = "П'ята"; } break; case 5: { s = "Шоста"; } break; }
                    xrowt = xrow; xcl1.WriteToCell(xrow, xcol, s + " пара"); xrow += 1;
                    xcl1.ws.Cells[xrow, xcol + 1].Font.Bold = true;
                    xcl1.ws.Cells[xrow, xcol + 1].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Cells[xrow, xcol + 1].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Cells[xrow, xcol + 1].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Cells[xrow, xcol + 1].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;

                    for (int z = 0; z < scheduleFree[x][y][0].Count; z++)
                    {
                        xcl1.WriteToCell(xrow, xcol, scheduleFree[x][y][0][z]);
                        xrow++;
                    }
                    xcl1.ws.Range[xcl1.ws.Cells[xrowt + 1, xcol + 1], xcl1.ws.Cells[xrow, xcol + 1]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range[xcl1.ws.Cells[xrowt + 1, xcol + 1], xcl1.ws.Cells[xrow, xcol + 1]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range[xcl1.ws.Cells[xrowt + 1, xcol + 1], xcl1.ws.Cells[xrow, xcol + 1]].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range[xcl1.ws.Cells[xrowt + 1, xcol + 1], xcl1.ws.Cells[xrow, xcol + 1]].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
                    xcol++;
                }
                xcol = 0;
                xcl1.ws.Cells[1, x + 1].Font.Bold = true;
                xcl1.ws.Cells[1, x + 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                xcl1.ws.Cells[1, x + 1].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                xcl1.ws.Cells[1, x + 1].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
                xcl1.ws.Cells[1, x + 1].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                xcl1.ws.Cells[1, x + 1].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
            }
            xcl1.ws.Columns.AutoFit();
            xcl1.SaveAs();
            xcl1.close();

            xcl1 = new Excel(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\Вільні аудиторії.xlsx"));
            xcol = 0;
            xrow = 1;
            xrowt = 0;

            cxz = new UInt16[5] { 0, 0, 0, 0, 0 };

            for (int y = 0; y < 5; y++)
            {
                for (int x = 0; x < 6; x++)
                {
                    if (scheduleFree[y][x][0].Count > cxz[y])
                        cxz[y] = Convert.ToUInt16(scheduleFree[y][x][0].Count);
                }
            }
            for (int x = 0; x < 5; x++)
            {
                xrow++;
                switch (x) { case 0: { xcl1.WriteToCell(0, 0, "Понеділок"); } break; case 1: { xcl1.WriteToCell(xrow, xcol, "Вівторок"); } break; case 2: { xcl1.WriteToCell(xrow, xcol, "Середа"); } break; case 3: { xcl1.WriteToCell(xrow, xcol, "Четвер"); } break; case 4: { xcl1.WriteToCell(xrow, xcol, "П'ятниця"); } break; }
                if (x == 0)
                {
                    xrow--;
                    xcl1.ws.Range["A" + (xrow).ToString(), "F" + (xrow).ToString()].Merge();
                    xcl1.ws.Range["A" + (xrow).ToString(), "F" + (xrow).ToString()].Font.Bold = true;
                    xcl1.ws.Range["A" + (xrow).ToString(), "F" + (xrow).ToString()].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range["A" + (xrow).ToString(), "F" + (xrow).ToString()].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range["A" + (xrow).ToString(), "F" + (xrow).ToString()].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range["A" + (xrow).ToString(), "F" + (xrow).ToString()].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range["A" + (xrow).ToString(), "F" + (xrow).ToString()].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                }
                else
                {
                    xcl1.ws.Range["A" + (xrow + 1).ToString(), "F" + (xrow + 1).ToString()].Merge();
                    xcl1.ws.Range["A" + (xrow + 1).ToString(), "F" + (xrow + 1).ToString()].Font.Bold = true;
                    xcl1.ws.Range["A" + (xrow + 1).ToString(), "F" + (xrow + 1).ToString()].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range["A" + (xrow + 1).ToString(), "F" + (xrow + 1).ToString()].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range["A" + (xrow + 1).ToString(), "F" + (xrow + 1).ToString()].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range["A" + (xrow + 1).ToString(), "F" + (xrow + 1).ToString()].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range["A" + (xrow + 1).ToString(), "F" + (xrow + 1).ToString()].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                }
                xrow = 1;
                for (int tx = 0; tx < x; tx++)
                {
                    xrow += cxz[tx] + 2;
                }
                for (int y = 0; y < 6; y++)
                {
                    xrow = 0;
                    for (int tx = 0; tx < x; tx++)
                    {
                        xrow += cxz[tx] + 2;
                    }
                    var s = "";
                    switch (y) { case 0: { s = "Перша"; } break; case 1: { s = "Друга"; } break; case 2: { s = "Третя"; } break; case 3: { s = "Четверта"; } break; case 4: { s = "П'ята"; } break; case 5: { s = "Шоста"; } break; }
                    xrowt = xrow; xcl1.WriteToCell(xrow, xcol, s + " пара"); xrow += 1;
                    xcl1.ws.Cells[xrow, xcol + 1].Font.Bold = true;
                    xcl1.ws.Cells[xrow, xcol + 1].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Cells[xrow, xcol + 1].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Cells[xrow, xcol + 1].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Cells[xrow, xcol + 1].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;

                    for (int z = 0; z < scheduleFree[x][y][1].Count; z++)
                    {
                        xcl1.WriteToCell(xrow, xcol, scheduleFree[x][y][1][z]);
                        xrow++;
                    }
                    xcl1.ws.Range[xcl1.ws.Cells[xrowt + 1, xcol + 1], xcl1.ws.Cells[xrow, xcol + 1]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range[xcl1.ws.Cells[xrowt + 1, xcol + 1], xcl1.ws.Cells[xrow, xcol + 1]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range[xcl1.ws.Cells[xrowt + 1, xcol + 1], xcl1.ws.Cells[xrow, xcol + 1]].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range[xcl1.ws.Cells[xrowt + 1, xcol + 1], xcl1.ws.Cells[xrow, xcol + 1]].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
                    xcol++;
                }
                xcol = 0;
                xcl1.ws.Cells[1, x + 1].Font.Bold = true;
                xcl1.ws.Cells[1, x + 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                xcl1.ws.Cells[1, x + 1].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                xcl1.ws.Cells[1, x + 1].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
                xcl1.ws.Cells[1, x + 1].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                xcl1.ws.Cells[1, x + 1].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
            }
            xcl1.ws.Columns.AutoFit();
            xcl1.SaveAs();
            xcl1.close();
        }

        public string getcolname(int inp)
        {
            var outp = "";
            while (inp > 0)
            {
                outp = Convert.ToChar((inp - 1) % 26 + 'A') + outp;
                inp = (inp - 1) / 26;
            }
            return outp;
        }

        // ?
        //private int func1(List<int> inp, int crs, int dy)
        //{
        //    var t = 5;
        //    while (inp[t] == 0)
        //        t--;
        //    return t;
        //}

        //private int func2(List<List<int>> inp)
        //{
        //    var outp = 0;

        //    for (int x = 0; x < inp.Count; x++)
        //    {
        //        for (int y = 0; y < inp[x].Count; y++)
        //        {
        //            outp += inp[x][y];
        //        }
        //    }
        //    return outp;
        //}
    }
}
