using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Schedule_Calculator_Pro
{
    public partial class Program : System.Windows.Window
    {
        // Оголошуємо всі змінні, які далі потрібно буде використовувати в інших класах:
        public static SortedDictionary<string, Don> don = new SortedDictionary<string, Don>();
        public static SortedDictionary<string, Group> group = new SortedDictionary<string, Group>();
        public static SortedDictionary<string, Subject> subject = new SortedDictionary<string, Subject>();
        public static SortedSet<string> audience = new SortedSet<string>();
        public static List<string> database = new List<string>();
        public static Schedule schedule = new Schedule();
        public static Thread SchedGenThread = new Thread(schedule.Start);
        public static Thread SavingThread = new Thread(CreateSettings);
        public static Thread PrimaryFileWorkThread = new Thread(SettingsHandle);
        public static CheckBox ChosenDay = null;

        public Program()
        {
            try
            {
                InitializeComponent();
                SchedGenThread.IsBackground = true;
                PrimaryFileWorkThread.IsBackground = true;
                PrimaryFileWorkThread.Priority = ThreadPriority.Highest;
                PrimaryFileWorkThread.Start();
                var temp = new Thread(CogAnimate);
                temp.IsBackground = true;
                temp.Start();
            }
            catch
            {
                MessageBox.Show("Виникла критична помилка. Перезапустіть програму та спробуйте ще раз. Якщо помилка повторюється - зв'яжіться із системним адміністратором.", "Упс...");
                Environment.Exit(111);
            }
        }


        #region Settings handlers
        public static void SettingsHandle()
        {
            //Перевіряємо, чи існує файл з налаштуваннями, відкриваємо його, якщо він є і створюємо у зворотньому випадку.
            if (File.Exists("База даних.xlsx"))
                GetSettings();
            else if (File.Exists("Дані.xlsx"))
            {
                CreateSettings();
            }
            else
            {
                MessageBox.Show("Створіть файл з початковими даними і назвіть його Дані.xlsx\n" +
                                "Типи даних по стовбцях: 1.Викладачі. 2.Групи 3.Предмети 4.Аудиторії");
                Environment.Exit(1);
            }
        }

        private static void GetSettings()
        {
            Excel excelTemp = new Excel(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\База даних.xlsx"), 1);
            var t = 0;
            while (excelTemp.BReadCell(t, 0))
            {
                var tmp = excelTemp.ReadCell(t, 0);       // преподаватель
                if (!don.ContainsKey(tmp))
                {
                    don.Add(tmp, new Don(tmp));
                }
                if (excelTemp.BReadCell(t, 1))
                {
                    don[tmp].relatedAud = excelTemp.ReadCell(t, 1);
                    audience.Add(excelTemp.ReadCell(t, 1));
                }
                if (excelTemp.BReadCell(t, 2))
                {
                    don[tmp].setExcludes(excelTemp.ReadCell(t, 2));
                }
                t++;
                while (excelTemp.BReadCell(t, 3))
                {
                    var tmp1 = excelTemp.ReadCell(t, 3);       // предмет
                    if (!subject.ContainsKey(tmp1))
                    {
                        subject.Add(tmp1, new Subject(tmp1));
                    }
                    if (!don[tmp].relatedSubjects.Contains(tmp1))
                        don[tmp].relatedSubjects.Add(tmp1);
                    if (excelTemp.BReadCell(t, 4))
                    {
                        subject[tmp1].relatedAud = excelTemp.ReadCell(t, 4);        // аудитория предмета
                        audience.Add(excelTemp.ReadCell(t, 4));
                    }
                    t++;
                    while (excelTemp.BReadCell(t, 5))
                    {
                        var tmp2 = excelTemp.ReadCell(t, 5);    // группа
                        if (!group.ContainsKey(tmp2))
                        {
                            group.Add(tmp2, new Group(tmp2));
                        }
                        if (!group[tmp2].relatedSubjects.ContainsKey(tmp1))
                        {
                            var tmp3 = "";
                            if (excelTemp.BReadCell(t, 6))
                                tmp3 = excelTemp.ReadCell(t, 6);            // кол-во пар в семестре
                            group[tmp2].SubjValEdit(tmp1, tmp, tmp3);
                        }
                        else if (!group[tmp2].relatedSubjects[tmp1].Contains(tmp))
                            group[tmp2].relatedSubjects[tmp1].Add(tmp);
                        if (excelTemp.BReadCell(t, 7))
                        {
                            group[tmp2].relatedAud = excelTemp.ReadCell(t, 7);
                            audience.Add(excelTemp.ReadCell(t, 7));
                        }
                        t++;
                    }
                }
            }

            var i = 0;

            while (excelTemp.BReadCell(i, 8))
            {
                var tmp = excelTemp.ReadCell(i, 8);
                if (!group.ContainsKey(tmp))
                {
                    group.Add(tmp, new Group(tmp));
                }
                i++;
            }
            i = 0;
            while (excelTemp.BReadCell(i, 9))
            {
                var tmp = excelTemp.ReadCell(i, 9);
                if (!subject.ContainsKey(tmp))
                {
                    subject.Add(tmp, new Subject(tmp));
                }
                i++;
            }
            i = 0;
            while (excelTemp.BReadCell(i, 10))
            {
                var tmp = excelTemp.ReadCell(i, 10);
                if (!group.ContainsKey(tmp))
                {
                    audience.Add(tmp);
                }
                i++;
            }

            while (excelTemp.BReadCell(i, 11))
            {
                var tmp = excelTemp.ReadCell(i, 11);

                {
                    group[tmp].StudyingWeeks = Convert.ToInt32(excelTemp.ReadCell(i, 12));
                }
                i++;
            }

            excelTemp.close();
            //MessageBox.Show("Завантаження початкових даних завершено.");
        }

        private static void CreateSettings()
        {
            if (!File.Exists("База даних.xlsx"))
            {
                Excel excelTemp = new Excel(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\Дані.xlsx"), 1);

                int it = 1, jt = 0;
                while (excelTemp.BReadCell(it, jt))
                {
                    don.Add(excelTemp.ReadCell(it, jt), new Don(excelTemp.ReadCell(it, jt)));
                    it++;
                }       // Преподаватель

                jt++; it = 1;
                while (excelTemp.BReadCell(it, jt))
                {
                    group.Add(excelTemp.ReadCell(it, jt), new Group(excelTemp.ReadCell(it, jt)));
                    it++;
                }       // Группа

                jt++; it = 1;
                while (excelTemp.BReadCell(it, jt))
                {
                    subject.Add(excelTemp.ReadCell(it, jt), new Subject(excelTemp.ReadCell(it, jt)));
                    it++;
                }       // Предмет

                jt++; it = 1;
                while (excelTemp.BReadCell(it, jt))
                {
                    audience.Add(excelTemp.ReadCell(it, jt));
                    it++;
                }       // Аудитория

                excelTemp.close();
            }
            Excel excelTemp1 = new Excel(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\База даних.xlsx"));
            List<List<string>> unwritten = new List<List<string>>();
            unwritten.Add(group.Keys.ToList()); unwritten.Add(subject.Keys.ToList()); unwritten.Add(audience.ToList());
            var t = 0;
            for (int j = 0; j < don.Count; j++)
            {
                var cdon = don.Keys.ToArray()[j];
                var cdonval = don.Values.ToArray()[j];
                excelTemp1.WriteToCell(t, 0, cdon);
                excelTemp1.WriteToCell(t, 1, cdonval.relatedAud);
                excelTemp1.WriteToCell(t, 2, cdonval.getExcludes());

                if (unwritten[2].Contains(cdonval.relatedAud))
                    unwritten[2].Remove(cdonval.relatedAud);

                t++;
                for (int k = 0; k < cdonval.relatedSubjects.Count; k++)
                {
                    var csubj = cdonval.relatedSubjects[k];
                    excelTemp1.WriteToCell(t, 3, csubj);
                    if (unwritten[1].Contains(csubj))
                        unwritten[1].Remove(csubj);

                    if (!subject.ContainsKey(csubj)) // ?
                        MessageBox.Show("Критична помилка: Цілісність даних порушена.");

                    excelTemp1.WriteToCell(t, 4, subject[csubj].relatedAud);

                    if (unwritten[2].Contains(subject[csubj].relatedAud))
                        unwritten[2].Remove(subject[csubj].relatedAud);
                    t++;
                    for (int g = 0; g < group.Count; g++)
                    {
                        var cgroup = group.Keys.ToArray()[g];
                        var cgroupval = group.Values.ToArray()[g];
                        if (cgroupval.relatedSubjects.ContainsKey(csubj) && cgroupval.relatedSubjects.Values.ToArray().Any(x => x.Contains(cdon)))
                        {
                            excelTemp1.WriteToCell(t, 5, cgroup);
                            if (unwritten[0].Contains(cgroup))
                                unwritten[0].Remove(cgroup);

                            excelTemp1.WriteToCell(t, 6, cgroupval.relatedSubjects[csubj][1].ToString());
                            excelTemp1.WriteToCell(t, 7, cgroupval.relatedAud);

                            if (unwritten[2].Contains(cgroupval.relatedAud))
                                unwritten[2].Remove(cgroupval.relatedAud);

                            t++;
                        }
                    }
                }
            }

            for (int i = 0; i < unwritten[0].Count; i++) // Сохранить не использованные группы
            {
                excelTemp1.WriteToCell(i, 8, unwritten[0][i]);
            }
            for (int i = 0; i < unwritten[1].Count; i++) // // Сохранить не использованные предметы
            {
                excelTemp1.WriteToCell(i, 9, unwritten[1][i]);
            }
            for (int i = 0; i < unwritten[2].Count; i++) // Сохранить не использованные аудитории
            {
                excelTemp1.WriteToCell(i, 10, unwritten[2][i]);
            }

            var temp1 = group.Keys.ToArray();
            var temp2 = group.Values.ToArray();
            for (int i = 0; i < temp1.Count(); i++) // Сохранить соотношение групп - учебных недель
            {
                excelTemp1.WriteToCell(i, 11, temp1[i]);
                excelTemp1.WriteToCell(i, 12, temp2[i].StudyingWeeks.ToString());
            }

            excelTemp1.ws.Columns.AutoFit();
            excelTemp1.SaveAs();
            excelTemp1.close();
            MessageBox.Show("Дані збережено.");
        }

        private void SearchHandle1(object sender, MouseEventArgs e)
        {
            //MessageBox.Show(e.GetPosition(grid).X.ToString() + " " + e.GetPosition(grid).Y.ToString());
            if (e.GetPosition(grid).Y >= 0 && e.GetPosition(grid).Y <= 30 && e.GetPosition(grid).X >= 30 && e.GetPosition(grid).X <= 915)
            {
                if (!(don.Keys.ToList().All(x => database.Contains(x)) && group.Keys.ToList().All(x => database.Contains(x)) && subject.Keys.ToList().All(x => database.Contains(x)) && audience.All(x => database.Contains(x))))
                {
                    database = new List<string>();
                    database.AddRange(don.Keys.ToArray());
                    database.AddRange(group.Keys.ToArray());
                    database.AddRange(subject.Keys.ToArray());
                    database.AddRange(audience.ToArray());
                }
                Search.ItemsSource = database.Where(x => x.ToLower().Contains(Search.Text.ToLower())).ToArray();

                if (!Search.IsDropDownOpen)
                    Search.IsDropDownOpen = true;
            }
        }
        private void SearchHandle2(object sender, EventArgs e)
        {
            if (!(don.Keys.ToList().All(x => database.Contains(x)) && group.Keys.ToList().All(x => database.Contains(x)) && subject.Keys.ToList().All(x => database.Contains(x)) && audience.All(x => database.Contains(x))))
            {
                database = new List<string>();
                database.AddRange(don.Keys.ToArray());
                database.AddRange(group.Keys.ToArray());
                database.AddRange(subject.Keys.ToArray());
                database.AddRange(audience.ToArray());
            }
            Search.ItemsSource = database.Where(x => x.ToLower().Contains(Search.Text.ToLower())).ToArray();

            if (!Search.IsDropDownOpen)
                Search.IsDropDownOpen = true;
        }
        #endregion

        #region Secondary funcs
        public static int DaynameToNum(string s)
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
            return n;
        }
        #endregion

        #region Animations
        private void MenuAnimate()
        {
            var n = -0.16;
            MenuX.Dispatcher.Invoke(delegate { MenuX.IsDefault = true; });
            if (MenuX.IsVisible)
            {
                Menu.Dispatcher.Invoke(delegate { Menu.Visibility = Visibility.Visible; });
                for (double x = .984; x >= 0; x -= .016)
                {
                    Menu.Dispatcher.Invoke(delegate { Menu.RenderTransform = new System.Windows.Media.RotateTransform(90 * x, 15, 15); Menu.Opacity = 1 - x; });
                    MenuX.Dispatcher.Invoke(delegate { MenuX.RenderTransform = new System.Windows.Media.RotateTransform(90 * x); MenuX.Opacity = x; });
                    MenuFreeDon.Dispatcher.Invoke(delegate { MenuFreeDon.Opacity = x; MenuFreeDon.Margin = new Thickness(MenuFreeDon.Margin.Left + n, MenuFreeDon.Margin.Top, MenuFreeDon.Margin.Right, MenuFreeDon.Margin.Bottom); });
                    MenuFreeAud.Dispatcher.Invoke(delegate { MenuFreeAud.Opacity = x; MenuFreeAud.Margin = new Thickness(MenuFreeAud.Margin.Left + n, MenuFreeAud.Margin.Top, MenuFreeAud.Margin.Right, MenuFreeAud.Margin.Bottom); });
                    MenuSchedule.Dispatcher.Invoke(delegate { MenuSchedule.Opacity = x; MenuSchedule.Margin = new Thickness(MenuSchedule.Margin.Left + n, MenuSchedule.Margin.Top, MenuSchedule.Margin.Right, MenuSchedule.Margin.Bottom); });
                    Thread.Sleep(2);
                }
                MenuX.Dispatcher.Invoke(delegate { MenuX.Visibility = Visibility.Hidden; });
                MenuFreeDon.Dispatcher.Invoke(delegate { MenuFreeDon.Visibility = Visibility.Hidden; });
                MenuFreeAud.Dispatcher.Invoke(delegate { MenuFreeAud.Visibility = Visibility.Hidden; });
                MenuSchedule.Dispatcher.Invoke(delegate { MenuSchedule.Visibility = Visibility.Hidden; });
            }
            else
            {
                MenuX.Dispatcher.Invoke(delegate { MenuX.Visibility = Visibility.Visible; });
                MenuFreeDon.Dispatcher.Invoke(delegate { MenuFreeDon.Visibility = Visibility.Visible; });
                MenuFreeAud.Dispatcher.Invoke(delegate { MenuFreeAud.Visibility = Visibility.Visible; });
                MenuSchedule.Dispatcher.Invoke(delegate { MenuSchedule.Visibility = Visibility.Visible; });
                for (double x = .016; x <= 1; x += .016)
                {
                    Menu.Dispatcher.Invoke(delegate { Menu.RenderTransform = new System.Windows.Media.RotateTransform(90 * x, 15, 15); Menu.Opacity = 1 - x; });
                    MenuX.Dispatcher.Invoke(delegate { MenuX.RenderTransform = new System.Windows.Media.RotateTransform(90 * x); MenuX.Opacity = x; });
                    MenuFreeDon.Dispatcher.Invoke(delegate { MenuFreeDon.Opacity = x; MenuFreeDon.Margin = new Thickness(MenuFreeDon.Margin.Left - n, MenuFreeDon.Margin.Top, MenuFreeDon.Margin.Right, MenuFreeDon.Margin.Bottom); });
                    MenuFreeAud.Dispatcher.Invoke(delegate { MenuFreeAud.Opacity = x; MenuFreeAud.Margin = new Thickness(MenuFreeAud.Margin.Left - n, MenuFreeAud.Margin.Top, MenuFreeAud.Margin.Right, MenuFreeAud.Margin.Bottom); });
                    MenuSchedule.Dispatcher.Invoke(delegate { MenuSchedule.Opacity = x; MenuSchedule.Margin = new Thickness(MenuSchedule.Margin.Left - n, MenuSchedule.Margin.Top, MenuSchedule.Margin.Right, MenuSchedule.Margin.Bottom); });
                    Thread.Sleep(2);
                }
                Menu.Dispatcher.Invoke(delegate { Menu.Visibility = Visibility.Hidden; });
            }
            MenuX.Dispatcher.Invoke(delegate { MenuX.IsDefault = false; });
        }

        private void DonAnimate()
        {
            if (donname.IsVisible)
            {
                for (double x = .98; x >= 0; x -= .02)
                {
                    donname.Dispatcher.Invoke(delegate { donname.Opacity = x; });
                    donrelatedsubjects.Dispatcher.Invoke(delegate { donrelatedsubjects.Opacity = x; });
                    deletedon.Dispatcher.Invoke(delegate { deletedon.Opacity = x; });
                    newdonrelsubj.Dispatcher.Invoke(delegate { newdonrelsubj.Opacity = x; });
                    donrelaud.Dispatcher.Invoke(delegate { donrelaud.Opacity = x; });
                    dond1.Dispatcher.Invoke(delegate { dond1.Opacity = x; });
                    dond2.Dispatcher.Invoke(delegate { dond2.Opacity = x; });
                    dond3.Dispatcher.Invoke(delegate { dond3.Opacity = x; });
                    dond4.Dispatcher.Invoke(delegate { dond4.Opacity = x; });
                    dond5.Dispatcher.Invoke(delegate { dond5.Opacity = x; });
                    Thread.Sleep(2);
                }
                donname.Dispatcher.Invoke(delegate { donname.Visibility = Visibility.Hidden; donname.Text = ""; });
                donrelatedsubjects.Dispatcher.Invoke(delegate { donrelatedsubjects.Visibility = Visibility.Hidden; donrelatedsubjects.ItemsSource = null; });
                deletedon.Dispatcher.Invoke(delegate { deletedon.Visibility = Visibility.Hidden; });
                newdonrelsubj.Dispatcher.Invoke(delegate { newdonrelsubj.Visibility = Visibility.Hidden; });
                donrelaud.Dispatcher.Invoke(delegate { donrelaud.Visibility = Visibility.Hidden; donrelaud.SelectedIndex = -1; });
                if (ChosenDay != null)
                    ChosenDay.Dispatcher.Invoke(delegate { ChosenDay.Content = ChosenDay.Content.ToString().Substring(0, ChosenDay.Content.ToString().Length - 1); ChosenDay = null; });
                dond1.Dispatcher.Invoke(delegate { dond1.Visibility = Visibility.Hidden; dond1.IsChecked = false; });
                dond2.Dispatcher.Invoke(delegate { dond2.Visibility = Visibility.Hidden; dond2.IsChecked = false; });
                dond3.Dispatcher.Invoke(delegate { dond3.Visibility = Visibility.Hidden; dond3.IsChecked = false; });
                dond4.Dispatcher.Invoke(delegate { dond4.Visibility = Visibility.Hidden; dond4.IsChecked = false; });
                dond5.Dispatcher.Invoke(delegate { dond5.Visibility = Visibility.Hidden; dond5.IsChecked = false; });
            }
            else
            {
                donname.Dispatcher.Invoke(delegate { donname.Visibility = Visibility.Visible; });
                donrelatedsubjects.Dispatcher.Invoke(delegate { donrelatedsubjects.Visibility = Visibility.Visible; });
                deletedon.Dispatcher.Invoke(delegate { deletedon.Visibility = Visibility.Visible; });
                newdonrelsubj.Dispatcher.Invoke(delegate { newdonrelsubj.Visibility = Visibility.Visible; });
                donrelaud.Dispatcher.Invoke(delegate { donrelaud.Visibility = Visibility.Visible; });
                dond1.Dispatcher.Invoke(delegate { dond1.Visibility = Visibility.Visible; });
                dond2.Dispatcher.Invoke(delegate { dond2.Visibility = Visibility.Visible; });
                dond3.Dispatcher.Invoke(delegate { dond3.Visibility = Visibility.Visible; });
                dond4.Dispatcher.Invoke(delegate { dond4.Visibility = Visibility.Visible; });
                dond5.Dispatcher.Invoke(delegate { dond5.Visibility = Visibility.Visible; });

                for (double x = .02; x < 1; x += .02)
                {
                    donname.Dispatcher.Invoke(delegate { donname.Opacity = x; });
                    donrelatedsubjects.Dispatcher.Invoke(delegate { donrelatedsubjects.Opacity = x; });
                    deletedon.Dispatcher.Invoke(delegate { deletedon.Opacity = x; });
                    newdonrelsubj.Dispatcher.Invoke(delegate { newdonrelsubj.Opacity = x; });
                    donrelaud.Dispatcher.Invoke(delegate { donrelaud.Opacity = x; });
                    dond1.Dispatcher.Invoke(delegate { dond1.Opacity = x; });
                    dond2.Dispatcher.Invoke(delegate { dond2.Opacity = x; });
                    dond3.Dispatcher.Invoke(delegate { dond3.Opacity = x; });
                    dond4.Dispatcher.Invoke(delegate { dond4.Opacity = x; });
                    dond5.Dispatcher.Invoke(delegate { dond5.Opacity = x; });
                    Thread.Sleep(2);
                }
            }
        }

        private void DonNewRelSubjAnimate()
        {
            if (donrelsubjcancel.IsVisible)
            {
                for (double x = .98; x >= 0; x -= .02)
                {
                    donrelsubjname.Dispatcher.Invoke(delegate { donrelsubjname.Opacity = x; });
                    donrelsubjok.Dispatcher.Invoke(delegate { donrelsubjok.Opacity = x; });
                    donrelsubjcancel.Dispatcher.Invoke(delegate { donrelsubjcancel.Opacity = x; });
                    Thread.Sleep(2);
                }
                donrelsubjname.Dispatcher.Invoke(delegate { donrelsubjname.Visibility = Visibility.Hidden; });
                donrelsubjok.Dispatcher.Invoke(delegate { donrelsubjok.Visibility = Visibility.Hidden; });
                donrelsubjcancel.Dispatcher.Invoke(delegate { donrelsubjcancel.Visibility = Visibility.Hidden; });
                donrelsubjname.Dispatcher.Invoke(delegate { donrelsubjname.Text = ""; });
            }
            else
            {
                donrelsubjname.Dispatcher.Invoke(delegate { donrelsubjname.Visibility = Visibility.Visible; });
                donrelsubjok.Dispatcher.Invoke(delegate { donrelsubjok.Visibility = Visibility.Visible; });
                donrelsubjcancel.Dispatcher.Invoke(delegate { donrelsubjcancel.Visibility = Visibility.Visible; });

                for (double x = .2; x < 1; x += .02)
                {
                    donrelsubjname.Dispatcher.Invoke(delegate { donrelsubjname.Opacity = x; });
                    donrelsubjok.Dispatcher.Invoke(delegate { donrelsubjok.Opacity = x; });
                    donrelsubjcancel.Dispatcher.Invoke(delegate { donrelsubjcancel.Opacity = x; });
                    Thread.Sleep(2);
                }
            }
        }

        private void DonRelSubjDelAnimate()
        {
            if (donrelsubjdel.IsVisible)
            {
                for (double x = .98; x >= 0; x -= .02)
                {
                    donrelsubjdel.Dispatcher.Invoke(delegate { donrelsubjdel.Opacity = x; });
                    Thread.Sleep(2);
                }
                donrelsubjdel.Dispatcher.Invoke(delegate { donrelsubjdel.Visibility = Visibility.Hidden; });
            }
            else
            {
                donrelsubjdel.Dispatcher.Invoke(delegate { donrelsubjdel.Visibility = Visibility.Visible; });
                for (double x = .02; x < 1; x += .02)
                {
                    donrelsubjdel.Dispatcher.Invoke(delegate { donrelsubjdel.Opacity = x; });
                    Thread.Sleep(2);
                }
            }
        }

        private void AudAnimate()
        {
            if (auddel.IsVisible)
            {
                for (double x = .98; x >= 0; x -= .02)
                {
                    auddel.Dispatcher.Invoke(delegate { auddel.Opacity = x; });
                    audname.Dispatcher.Invoke(delegate { audname.Opacity = x; });
                    audsave.Dispatcher.Invoke(delegate { audsave.Opacity = x; });
                    Thread.Sleep(2);
                }
                auddel.Dispatcher.Invoke(delegate { auddel.Visibility = Visibility.Hidden; });
                audname.Dispatcher.Invoke(delegate { audname.Visibility = Visibility.Hidden; audname.Text = ""; });
                audsave.Dispatcher.Invoke(delegate { audsave.Visibility = Visibility.Hidden; });
            }
            else
            {
                audname.Dispatcher.Invoke(delegate { audname.Visibility = Visibility.Visible; });
                auddel.Dispatcher.Invoke(delegate { auddel.Visibility = Visibility.Visible; });
                audsave.Dispatcher.Invoke(delegate { audsave.Visibility = Visibility.Visible; });
                for (double x = .02; x < 1; x += .02)
                {
                    auddel.Dispatcher.Invoke(delegate { auddel.Opacity = x; });
                    audname.Dispatcher.Invoke(delegate { audname.Opacity = x; });
                    audsave.Dispatcher.Invoke(delegate { audsave.Opacity = x; });
                    Thread.Sleep(2);
                }
            }
        }

        private void GroupAnimate()
        {
            if (groupname.IsVisible)
            {
                for (double x = .98; x >= 0; x -= .02)
                {
                    groupname.Dispatcher.Invoke(delegate { groupname.Opacity = x; });
                    grouprelatedinfo.Dispatcher.Invoke(delegate { grouprelatedinfo.Opacity = x; });
                    deletegroup.Dispatcher.Invoke(delegate { deletegroup.Opacity = x; });
                    newgrouprelsubj.Dispatcher.Invoke(delegate { newgrouprelsubj.Opacity = x; });
                    grouprelaud.Dispatcher.Invoke(delegate { grouprelaud.Opacity = x; });
                    groupstudyingweeks.Dispatcher.Invoke(delegate { groupstudyingweeks.Opacity = x; });
                    groupsavechanges.Dispatcher.Invoke(delegate { groupsavechanges.Opacity = x; });
                    Thread.Sleep(2);
                }
                groupname.Dispatcher.Invoke(delegate { groupname.Visibility = Visibility.Hidden; groupname.Text = ""; });
                grouprelatedinfo.Dispatcher.Invoke(delegate { grouprelatedinfo.Visibility = Visibility.Hidden; });
                deletegroup.Dispatcher.Invoke(delegate { deletegroup.Visibility = Visibility.Hidden; });
                newgrouprelsubj.Dispatcher.Invoke(delegate { newgrouprelsubj.Visibility = Visibility.Hidden; });
                grouprelaud.Dispatcher.Invoke(delegate { grouprelaud.Visibility = Visibility.Hidden; grouprelaud.Text = ""; });
                groupstudyingweeks.Dispatcher.Invoke(delegate { groupstudyingweeks.Visibility = Visibility.Hidden; groupstudyingweeks.Text = ""; });
                groupsavechanges.Dispatcher.Invoke(delegate { groupsavechanges.Visibility = Visibility.Hidden; });
            }
            else
            {
                groupname.Dispatcher.Invoke(delegate { groupname.Visibility = Visibility.Visible; });
                grouprelatedinfo.Dispatcher.Invoke(delegate { grouprelatedinfo.Visibility = Visibility.Visible; });
                deletegroup.Dispatcher.Invoke(delegate { deletegroup.Visibility = Visibility.Visible; });
                newgrouprelsubj.Dispatcher.Invoke(delegate { newgrouprelsubj.Visibility = Visibility.Visible; });
                grouprelaud.Dispatcher.Invoke(delegate { grouprelaud.Visibility = Visibility.Visible; });
                groupstudyingweeks.Dispatcher.Invoke(delegate { groupstudyingweeks.Visibility = Visibility.Visible; });
                groupsavechanges.Dispatcher.Invoke(delegate { groupsavechanges.Visibility = Visibility.Visible; });

                for (double x = .2; x < 1; x += .02)
                {
                    groupname.Dispatcher.Invoke(delegate { groupname.Opacity = x; });
                    grouprelatedinfo.Dispatcher.Invoke(delegate { grouprelatedinfo.Opacity = x; });
                    deletegroup.Dispatcher.Invoke(delegate { deletegroup.Opacity = x; });
                    newgrouprelsubj.Dispatcher.Invoke(delegate { newgrouprelsubj.Opacity = x; });
                    grouprelaud.Dispatcher.Invoke(delegate { grouprelaud.Opacity = x; });
                    groupstudyingweeks.Dispatcher.Invoke(delegate { groupstudyingweeks.Opacity = x; });
                    groupsavechanges.Dispatcher.Invoke(delegate { groupsavechanges.Opacity = x; });
                    Thread.Sleep(2);
                }
            }
        }

        private void GroupNewRelSubjAnimate()
        {
            if (editsubjname.IsVisible)
            {
                for (double x = .98; x >= 0; x -= .02)
                {
                    editsubjname.Dispatcher.Invoke(delegate { editsubjname.Opacity = x; });
                    editdonname.Dispatcher.Invoke(delegate { editdonname.Opacity = x; });
                    editcoupleahalf.Dispatcher.Invoke(delegate { editcoupleahalf.Opacity = x; });
                    editseconddonname.Dispatcher.Invoke(delegate { editseconddonname.Opacity = x; });
                    grouprelsubjok.Dispatcher.Invoke(delegate { grouprelsubjok.Opacity = x; });
                    grouprelsubjcancel.Dispatcher.Invoke(delegate { grouprelsubjcancel.Opacity = x; });
                    Thread.Sleep(2);
                }
                editsubjname.Dispatcher.Invoke(delegate { editsubjname.Visibility = Visibility.Hidden; editsubjname.Text = ""; });
                editdonname.Dispatcher.Invoke(delegate { editdonname.Visibility = Visibility.Hidden; editdonname.Text = ""; });
                editcoupleahalf.Dispatcher.Invoke(delegate { editcoupleahalf.Visibility = Visibility.Hidden; editcoupleahalf.Text = ""; });
                editseconddonname.Dispatcher.Invoke(delegate { editseconddonname.Visibility = Visibility.Hidden; editseconddonname.Text = ""; });
                grouprelsubjok.Dispatcher.Invoke(delegate { grouprelsubjok.Visibility = Visibility.Hidden; });
                grouprelsubjcancel.Dispatcher.Invoke(delegate { grouprelsubjcancel.Visibility = Visibility.Hidden; });
            }
            else
            {
                editsubjname.Dispatcher.Invoke(delegate { editsubjname.Visibility = Visibility.Visible; });
                editdonname.Dispatcher.Invoke(delegate { editdonname.Visibility = Visibility.Visible; });
                editcoupleahalf.Dispatcher.Invoke(delegate { editcoupleahalf.Visibility = Visibility.Visible; });
                editseconddonname.Dispatcher.Invoke(delegate { editseconddonname.Visibility = Visibility.Visible; });
                grouprelsubjok.Dispatcher.Invoke(delegate { grouprelsubjok.Visibility = Visibility.Visible; });
                grouprelsubjcancel.Dispatcher.Invoke(delegate { grouprelsubjcancel.Visibility = Visibility.Visible; });
                for (double x = .2; x < 1; x += .02)
                {
                    editsubjname.Dispatcher.Invoke(delegate { editsubjname.Opacity = x; });
                    editdonname.Dispatcher.Invoke(delegate { editdonname.Opacity = x; });
                    editcoupleahalf.Dispatcher.Invoke(delegate { editcoupleahalf.Opacity = x; });
                    editseconddonname.Dispatcher.Invoke(delegate { editseconddonname.Opacity = x; });
                    grouprelsubjok.Dispatcher.Invoke(delegate { grouprelsubjok.Opacity = x; });
                    grouprelsubjcancel.Dispatcher.Invoke(delegate { grouprelsubjcancel.Opacity = x; });
                    Thread.Sleep(2);
                }

            }

        }

        private void GroupRelSubjDelAnimate()
        {
            if (grouprelsubjdel.IsVisible)
            {
                for (double x = .98; x >= 0; x -= .02)
                {
                    grouprelsubjdel.Dispatcher.Invoke(delegate { grouprelsubjdel.Opacity = x; });
                    Thread.Sleep(2);
                }
                grouprelsubjdel.Dispatcher.Invoke(delegate { grouprelsubjdel.Visibility = Visibility.Hidden; });
            }
            else
            {
                grouprelsubjdel.Dispatcher.Invoke(delegate { grouprelsubjdel.Visibility = Visibility.Visible; });
                for (double x = .02; x < 1; x += .02)
                {
                    grouprelsubjdel.Dispatcher.Invoke(delegate { grouprelsubjdel.Opacity = x; });
                    Thread.Sleep(2);
                }
            }
        }

        private void SubjectAnimate()
        {
            if (subjectname.IsVisible)
            {
                for (double x = .98; x >= 0; x -= .02)
                {
                    subjectname.Dispatcher.Invoke(delegate { subjectname.Opacity = x; });
                    subjecttext.Dispatcher.Invoke(delegate { subjecttext.Opacity = x; });
                    subjectrelaud.Dispatcher.Invoke(delegate { subjectrelaud.Opacity = x; });
                    audtext.Dispatcher.Invoke(delegate { audtext.Opacity = x; });
                    subjsavechanges.Dispatcher.Invoke(delegate { subjsavechanges.Opacity = x; });
                    deletesubject.Dispatcher.Invoke(delegate { deletesubject.Opacity = x; });
                    Thread.Sleep(2);
                }
                subjectname.Dispatcher.Invoke(delegate { subjectname.Visibility = Visibility.Hidden; subjectname.Text = ""; });
                subjecttext.Dispatcher.Invoke(delegate { subjecttext.Visibility = Visibility.Hidden; subjecttext.Text = ""; });
                subjectrelaud.Dispatcher.Invoke(delegate { subjectrelaud.Visibility = Visibility.Hidden; subjectrelaud.Text = ""; });
                audtext.Dispatcher.Invoke(delegate { audtext.Visibility = Visibility.Hidden; audtext.Text = ""; });
                subjsavechanges.Dispatcher.Invoke(delegate { subjsavechanges.Visibility = Visibility.Hidden; });
                deletesubject.Dispatcher.Invoke(delegate { deletesubject.Visibility = Visibility.Hidden; });
            }
            else
            {
                subjectname.Dispatcher.Invoke(delegate { subjectname.Visibility = Visibility.Visible; });
                subjecttext.Dispatcher.Invoke(delegate { subjecttext.Visibility = Visibility.Visible; });
                subjectrelaud.Dispatcher.Invoke(delegate { subjectrelaud.Visibility = Visibility.Visible; });
                audtext.Dispatcher.Invoke(delegate { audtext.Visibility = Visibility.Visible; });
                subjsavechanges.Dispatcher.Invoke(delegate { subjsavechanges.Visibility = Visibility.Visible; });
                deletesubject.Dispatcher.Invoke(delegate { deletesubject.Visibility = Visibility.Visible; });
                for (double x = .2; x < 1; x += .02)
                {
                    subjectname.Dispatcher.Invoke(delegate { subjectname.Opacity = x; });
                    subjecttext.Dispatcher.Invoke(delegate { subjecttext.Opacity = x; });
                    subjectrelaud.Dispatcher.Invoke(delegate { subjectrelaud.Opacity = x; });
                    audtext.Dispatcher.Invoke(delegate { audtext.Opacity = x; });
                    subjsavechanges.Dispatcher.Invoke(delegate { subjsavechanges.Opacity = x; });
                    deletesubject.Dispatcher.Invoke(delegate { deletesubject.Opacity = x; });

                    Thread.Sleep(2);
                }
            }
        }

        public void CogAnimate()
        {
            while (true)
            {
                if (SchedGenThread.IsAlive || PrimaryFileWorkThread.IsAlive || SavingThread.IsAlive)
                {
                    rotcog.Dispatcher.Invoke(delegate { rotcog.Visibility = Visibility.Visible; });
                    for (double x = .02; x <= 1; x += .02)
                    {
                        rotcog.Dispatcher.Invoke(delegate { rotcog.Opacity = x; });
                        Thread.Sleep(4);
                    }
                    var angle = 0;
                    while (SchedGenThread.IsAlive || PrimaryFileWorkThread.IsAlive || SavingThread.IsAlive)
                    {
                        rotcog.Dispatcher.Invoke(delegate { rotcog.RenderTransform = new System.Windows.Media.RotateTransform(angle, 50, 50); });
                        angle += 3;
                        if (angle > 360)
                            angle = 0;
                        Thread.Sleep(6);
                    }
                    for (double x = .98; x >= 0; x -= .02)
                    {
                        rotcog.Dispatcher.Invoke(delegate { rotcog.Opacity = x; });
                        Thread.Sleep(4);
                    }
                    rotcog.Dispatcher.Invoke(delegate { rotcog.Visibility = Visibility.Hidden; });
                    Thread.Sleep(250);
                }
            }
        }

        private void CoupleExcAnimate()
        {
            if (donc1.IsVisible)
            {
                for (double x = .98; x >= 0; x -= .02)
                {
                    donc1.Dispatcher.Invoke(delegate { donc1.Opacity = x; });
                    donc2.Dispatcher.Invoke(delegate { donc2.Opacity = x; });
                    donc3.Dispatcher.Invoke(delegate { donc3.Opacity = x; });
                    donc4.Dispatcher.Invoke(delegate { donc4.Opacity = x; });
                    donc5.Dispatcher.Invoke(delegate { donc5.Opacity = x; });
                    donc6.Dispatcher.Invoke(delegate { donc6.Opacity = x; });
                    Thread.Sleep(2);
                }
                donc1.Dispatcher.Invoke(delegate { donc1.Visibility = Visibility.Hidden; donc1.IsChecked = false; });
                donc2.Dispatcher.Invoke(delegate { donc2.Visibility = Visibility.Hidden; donc2.IsChecked = false; });
                donc3.Dispatcher.Invoke(delegate { donc3.Visibility = Visibility.Hidden; donc3.IsChecked = false; });
                donc4.Dispatcher.Invoke(delegate { donc4.Visibility = Visibility.Hidden; donc4.IsChecked = false; });
                donc5.Dispatcher.Invoke(delegate { donc5.Visibility = Visibility.Hidden; donc5.IsChecked = false; });
                donc6.Dispatcher.Invoke(delegate { donc6.Visibility = Visibility.Hidden; donc6.IsChecked = false; });
            }
            else
            {
                donc1.Dispatcher.Invoke(delegate { donc1.Visibility = Visibility.Visible; });
                donc2.Dispatcher.Invoke(delegate { donc2.Visibility = Visibility.Visible; });
                donc3.Dispatcher.Invoke(delegate { donc3.Visibility = Visibility.Visible; });
                donc4.Dispatcher.Invoke(delegate { donc4.Visibility = Visibility.Visible; });
                donc5.Dispatcher.Invoke(delegate { donc5.Visibility = Visibility.Visible; });
                donc6.Dispatcher.Invoke(delegate { donc6.Visibility = Visibility.Visible; });

                for (double x = .02; x < 1; x += .02)
                {
                    donc1.Dispatcher.Invoke(delegate { donc1.Opacity = x; });
                    donc2.Dispatcher.Invoke(delegate { donc2.Opacity = x; });
                    donc3.Dispatcher.Invoke(delegate { donc3.Opacity = x; });
                    donc4.Dispatcher.Invoke(delegate { donc4.Opacity = x; });
                    donc5.Dispatcher.Invoke(delegate { donc5.Opacity = x; });
                    donc6.Dispatcher.Invoke(delegate { donc6.Opacity = x; });
                    Thread.Sleep(2);
                }
            }
        }
        #endregion


        #region Callbacks

        #region Click callbacks
        private void Menu_Click(object sender, RoutedEventArgs e)
        {
            if (!MenuX.IsDefault)
            {
                var Anim = new Thread(MenuAnimate);
                Anim.IsBackground = true;
                Anim.Start();
            }
        }

        private void Grid_Click(object sender, MouseButtonEventArgs e)
        {
            if (MenuX.IsVisible && !MenuX.IsDefault)
            {
                var temp = new Thread(MenuAnimate);
                temp.IsBackground = true;
                temp.Start();
            }
            if (donrelsubjname.IsVisible)
            {
                var temp = new Thread(DonNewRelSubjAnimate);
                temp.IsBackground = true;
                temp.Start();
            }
            var tmp = Mouse.GetPosition(System.Windows.Application.Current.MainWindow);
            if (editdonname.IsVisible && !(tmp.X >= 213 && tmp.X < 741 && tmp.Y >= 219 && tmp.Y < 393))
            {
                var temp = new Thread(GroupNewRelSubjAnimate);
                temp.IsBackground = true;
                temp.Start();
            }
            if (donrelsubjdel.IsVisible && !(tmp.X >= 300 && tmp.X < 628 && tmp.Y >= 267 && tmp.Y < 345))
            {
                var temp = new Thread(DonRelSubjDelAnimate);
                temp.IsBackground = true;
                temp.Start();
            }
            if (grouprelsubjdel.IsVisible && !(tmp.X >= 213 && tmp.X < 741 && tmp.Y >= 219 && tmp.Y < 393))
            {
                var temp = new Thread(GroupRelSubjDelAnimate);
                temp.IsBackground = true;
                temp.Start();
            }
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            Search.SelectedIndex = -1;
            if (!(don.Keys.ToList().All(x => database.Contains(x)) && group.Keys.ToList().All(x => database.Contains(x)) && subject.Keys.ToList().All(x => database.Contains(x)) && audience.ToList().All(x => database.Contains(x))))
            {
                database.RemoveRange(0, database.Count);
                database.AddRange(don.Keys.ToArray());
                database.AddRange(group.Keys.ToArray());
                database.AddRange(subject.Keys.ToArray());
                database.AddRange(audience.ToArray());
            }
            Search.ItemsSource = database;
            if (donname.IsVisible)
            {
                var temp = new Thread(DonAnimate);
                temp.IsBackground = true;
                temp.Start();
            }
            if (donrelsubjname.IsVisible)
            {
                var temp = new Thread(DonNewRelSubjAnimate);
                temp.IsBackground = true;
                temp.Start();
            }
            if (donrelsubjdel.IsVisible)
            {
                var temp = new Thread(DonRelSubjDelAnimate);
                temp.IsBackground = true;
                temp.Start();
            }
            if (groupname.IsVisible)
            {
                var temp = new Thread(GroupAnimate);
                temp.IsBackground = true;
                temp.Start();
            }
            if (editsubjname.IsVisible)
            {
                var temp = new Thread(GroupNewRelSubjAnimate);
                temp.IsBackground = true;
                temp.Start();
            }
            if (grouprelsubjdel.IsVisible)
            {
                var temp = new Thread(GroupRelSubjDelAnimate);
                temp.IsBackground = true;
                temp.Start();
            }
            if (subjectname.IsVisible)
            {
                var temp = new Thread(SubjectAnimate);
                temp.IsBackground = true;
                temp.Start();
            }
            if (donc1.IsVisible)
            {
                var temp = new Thread(CoupleExcAnimate);
                temp.IsBackground = true;
                temp.Start();
            }

        }

        private void save_Click(object sender, RoutedEventArgs e)
        {
            Clear_Click(sender, e);
            if (SavingThread.IsAlive)
                MessageBox.Show("Збереження вже відбувається...");
            else if (SchedGenThread.IsAlive)
                MessageBox.Show("Збереження неможливо, тому що відбувається створення розкладу.");
            else
            {
                SavingThread = new Thread(CreateSettings);
                SavingThread.IsBackground = true;
                SavingThread.Start();
            }
        }

        private void newdonrelsubj_Click(object sender, RoutedEventArgs e)
        {
            if (!donrelsubjname.IsVisible)
            {
                var temp = new Thread(DonNewRelSubjAnimate);
                temp.IsBackground = true;
                temp.Start();
            }
        }

        private void donrelsubjcancel_Click(object sender, RoutedEventArgs e)
        {
            var temp = new Thread(DonNewRelSubjAnimate);
            temp.IsBackground = true;
            temp.Start();
        }

        private void donrelsubjok_Click(object sender, RoutedEventArgs e)
        {
            if (!don[Search.SelectedItem.ToString()].relatedSubjects.Contains(donrelsubjname.Text))
            {
                donrelatedsubjects.ItemsSource = null;
                don[Search.SelectedItem.ToString()].relatedSubjects.Add(donrelsubjname.Text);
                if (!subject.ContainsKey(donrelsubjname.Text))
                    subject.Add(donrelsubjname.Text, new Subject(donrelsubjname.Text));
                donrelatedsubjects.ItemsSource = don[Search.SelectedItem.ToString()].relatedSubjects;
            }
            var temp = new Thread(DonNewRelSubjAnimate);
            temp.IsBackground = true;
            temp.Start();
        }

        private void deletedon_Click(object sender, RoutedEventArgs e)
        {
            don.Remove(Search.SelectedItem.ToString());
            foreach (var key in group.Keys)
            {
                for (int x = 0; x < group[key].relatedSubjects.Count; x++)
                    if (group[key].relatedSubjects[group[key].relatedSubjects.Keys.ToArray()[x]].Contains(Search.SelectedItem.ToString()))
                    {
                        group[key].relatedSubjects.Remove(group[key].relatedSubjects.Keys.ToArray()[x]);
                        x--;
                    }

            }
            Search.SelectedIndex = -1;
            database.RemoveRange(0, database.Count);
            database.AddRange(don.Keys.ToArray());
            database.AddRange(group.Keys.ToArray());
            database.AddRange(subject.Keys.ToArray());
            database.AddRange(audience.ToArray());
            Search.ItemsSource = database;
            if (donrelsubjname.IsVisible)
            {
                var temp1 = new Thread(DonNewRelSubjAnimate);
                temp1.IsBackground = true;
                temp1.Start();
            }
            var temp = new Thread(DonAnimate);
            temp.IsBackground = true;
            temp.Start();

        }

        private void donrelsubjdel_Click(object sender, RoutedEventArgs e)
        {
            don[Search.SelectedItem.ToString()].relatedSubjects.Remove(donrelatedsubjects.SelectedItem.ToString());
            donrelatedsubjects.ItemsSource = null;
            donrelatedsubjects.ItemsSource = don[Search.SelectedItem.ToString()].relatedSubjects;
            var temp = new Thread(DonRelSubjDelAnimate);
            temp.IsBackground = true;
            temp.Start();
        }

        private void schedulegenerate_Click(object sender, RoutedEventArgs e)
        {
            if (SchedGenThread.IsAlive)
                MessageBox.Show("Розклад вже генерується...");
            else if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\SCP\\Schedule.xlsx"))
                MessageBox.Show("Розклад вже створено.");
            else if (PrimaryFileWorkThread.IsAlive || SavingThread.IsAlive)
                MessageBox.Show("Зараз неможливо створити розклад тому, що відбувається зчитування чи створення файлу.");
            else
            {
                SchedGenThread = new Thread(schedule.Start);
                SchedGenThread.IsBackground = true;
                SchedGenThread.Start();
            }

        }

        private void auddel_Click(object sender, RoutedEventArgs e)
        {
            audience.Remove(Search.SelectedItem.ToString());
            for (int x = 0; x < group.Count; x++)
                if (group.Values.ToArray()[x].relatedAud == Search.SelectedItem.ToString())
                    group[group.Keys.ToArray()[x]].relatedAud = "";
            for (int x = 0; x < subject.Count; x++)
                if (subject.Values.ToArray()[x].relatedAud == Search.SelectedItem.ToString())
                    subject[subject.Keys.ToArray()[x]].relatedAud = "";
            for (int x = 0; x < don.Count; x++)
                if (don.Values.ToArray()[x].relatedAud == Search.SelectedItem.ToString())
                    don[don.Keys.ToArray()[x]].relatedAud = "";
            Search.SelectedIndex = -1;
            database.RemoveRange(0, database.Count);
            database.AddRange(don.Keys.ToArray());
            database.AddRange(group.Keys.ToArray());
            database.AddRange(subject.Keys.ToArray());
            database.AddRange(audience.ToArray());
            Search.ItemsSource = database;
            var temp = new Thread(AudAnimate);
            temp.IsBackground = true;
            temp.Start();
        }

        private void MenuSchedule_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\Розклад 1 курс.xlsx"));
            System.Diagnostics.Process.Start(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\Розклад 2 курс.xlsx"));
            System.Diagnostics.Process.Start(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\Розклад 3 курс.xlsx"));
            System.Diagnostics.Process.Start(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\Розклад 4 курс.xlsx"));
        }

        private void deletegroup_Click(object sender, RoutedEventArgs e)
        {
            var temp = new Thread(GroupAnimate);
            temp.IsBackground = true;
            temp.Start();
            if (grouprelsubjdel.IsVisible)
            {
                var temp1 = new Thread(GroupRelSubjDelAnimate);
                temp1.IsBackground = true;
                temp1.Start();
            }
            if (editdonname.IsVisible)
            {
                var temp1 = new Thread(GroupNewRelSubjAnimate);
                temp1.IsBackground = true;
                temp1.Start();
            }
            var temp2 = new Thread(GroupAnimate);
            temp2.IsBackground = true;
            temp2.Start();
            group.Remove(Search.SelectedItem.ToString());
            Search.SelectedIndex = -1;
        }

        private void grouprelsubjdel_Click(object sender, RoutedEventArgs e)
        {
            group[Search.SelectedItem.ToString()].relatedSubjects.Remove(group[Search.SelectedItem.ToString()].relatedSubjects.Keys.ToArray()[grouprelatedinfo.SelectedIndex]);
            grouprelatedinfo.ItemsSource = null;
            var temp1 = group[Search.SelectedItem.ToString()].relatedSubjects;
            List<List<string>> infoset = new List<List<string>>();
            for (int x = 0; x < temp1.Count; x++)
            {
                infoset.Add(new List<string>());
                infoset[x].AddRange(new List<string>() { "", "", "", "" });
                infoset[x][0] = temp1.Keys.ToArray()[x];
                infoset[x][1] = temp1.Values.ToArray()[x][0];
                infoset[x][2] = temp1.Values.ToArray()[x][1];
                if (temp1.Values.ToArray()[x].Count == 3)
                    infoset[x][3] = temp1.Values.ToArray()[x][2];
            }
            grouprelatedinfo.ItemsSource = infoset;
            var temp = new Thread(GroupRelSubjDelAnimate);
            temp.IsBackground = true;
            temp.Start();
            var temp2 = new Thread(GroupNewRelSubjAnimate);
            temp2.IsBackground = true;
            temp2.Start();
        }

        private void groupsavechanges_Click(object sender, RoutedEventArgs e)
        {
            group[Search.SelectedItem.ToString()].StudyingWeeks = Convert.ToInt32(groupstudyingweeks.Text);
            group[Search.SelectedItem.ToString()].relatedAud = grouprelaud.Text;
            if (Search.SelectedItem.ToString() != groupname.Text)
            {
                group[Search.SelectedItem.ToString()].groupName = groupname.Text;
                group.Add(groupname.Text, group[Search.SelectedItem.ToString()]);
                group.Remove(Search.SelectedItem.ToString());
            }
            var temp = new Thread(GroupAnimate);
            temp.IsBackground = true;
            temp.Start();
        }

        private void grouprelsubjcancel_Click(object sender, RoutedEventArgs e)
        {
            var temp = new Thread(GroupNewRelSubjAnimate);
            temp.IsBackground = true;
            temp.Start();
        }

        private void grouprelsubjok_Click(object sender, RoutedEventArgs e)
        {
            if (grouprelatedinfo.SelectedIndex != -1)
            {
                var temp = group[Search.SelectedItem.ToString()].relatedSubjects;
                temp.Values.ToArray()[grouprelatedinfo.SelectedIndex][0] = editdonname.Text;
                temp.Values.ToArray()[grouprelatedinfo.SelectedIndex][1] = editcoupleahalf.Text;
                if (temp.Values.ToArray().Length == 3)
                    temp.Values.ToArray()[grouprelatedinfo.SelectedIndex][2] = editseconddonname.Text;
                else if (editseconddonname.Text != "")
                    temp.Values.ToArray()[grouprelatedinfo.SelectedIndex].Add(editseconddonname.Text);
                if (group[Search.SelectedItem.ToString()].relatedSubjects.Keys.ToArray()[grouprelatedinfo.SelectedIndex] != editsubjname.Text)
                {
                    group[Search.SelectedItem.ToString()].relatedSubjects.Add(editsubjname.Text, new List<string>(group[Search.SelectedItem.ToString()].relatedSubjects.Values.ToArray()[grouprelatedinfo.SelectedIndex]));
                    group[Search.SelectedItem.ToString()].relatedSubjects.Remove(group[Search.SelectedItem.ToString()].relatedSubjects.Keys.ToArray()[grouprelatedinfo.SelectedIndex]);
                }
            }
            else
            {
                if (don.ContainsKey(editdonname.Text))
                {
                    if (!don[editdonname.Text].relatedSubjects.Contains(editsubjname.Text))
                        don[editdonname.Text].relatedSubjects.Add(editsubjname.Text);
                }
                else
                {
                    don.Add(editdonname.Text, new Don(editdonname.Text));
                    don[editdonname.Text].relatedSubjects.Add(editsubjname.Text);
                }
                if (don.ContainsKey(editseconddonname.Text))
                {
                    if (!don[editseconddonname.Text].relatedSubjects.Contains(editseconddonname.Text))
                        don[editseconddonname.Text].relatedSubjects.Add(editseconddonname.Text);
                }
                else
                {
                    don.Add(editseconddonname.Text, new Don(editseconddonname.Text));
                    don[editseconddonname.Text].relatedSubjects.Add(editsubjname.Text);
                }
                group[Search.SelectedItem.ToString()].relatedSubjects.Add(editsubjname.Text, new List<string>() { editdonname.Text, editcoupleahalf.Text, editseconddonname.Text });
            }
            if (!subject.ContainsKey(editsubjname.Text))
                subject.Add(editsubjname.Text, new Subject(editsubjname.Text));
            grouprelatedinfo.SelectedIndex = -1;
            grouprelatedinfo.ItemsSource = null;
            var temp1 = group[Search.SelectedItem.ToString()].relatedSubjects;
            List<List<string>> infoset = new List<List<string>>();
            for (int x = 0; x < temp1.Count; x++)
            {
                infoset.Add(new List<string>());
                infoset[x].AddRange(new List<string>() { "", "", "", "" });
                infoset[x][0] = temp1.Keys.ToArray()[x];
                infoset[x][1] = temp1.Values.ToArray()[x][0];
                infoset[x][2] = temp1.Values.ToArray()[x][1];
                if (temp1.Values.ToArray()[x].Count == 3)
                    infoset[x][3] = temp1.Values.ToArray()[x][2];
            }
            grouprelatedinfo.ItemsSource = infoset;
            var temp2 = new Thread(GroupNewRelSubjAnimate);
            temp2.IsBackground = true;
            temp2.Start();
        }

        private void newgrouprelsubj_Click(object sender, RoutedEventArgs e)
        {
            grouprelatedinfo.SelectedIndex = -1;
            if (!editdonname.IsVisible)
            {
                var temp = new Thread(GroupNewRelSubjAnimate);
                temp.IsBackground = true;
                temp.Start();
            }
            else
            {
                editdonname.Text = "";
                editsubjname.Text = "";
                editcoupleahalf.Text = "";
                editseconddonname.Text = "";
            }
        }

        private void subjsavechanges_Click(object sender, RoutedEventArgs e)
        {
            subject[Search.Text].subjectName = subjectname.Text;
            subject[Search.Text].relatedAud = subjectrelaud.Text;
            var temp = new Thread(SubjectAnimate);
            temp.IsBackground = true;
            temp.Start();
        }

        private void deletesubject_Click(object sender, RoutedEventArgs e)
        {
            for (int x = 0; x < group.Count; x++)
            {
                if (group.Values.ToArray()[x].relatedSubjects.ContainsKey(subjectname.Text))
                    group[group.Keys.ToArray()[x]].relatedSubjects.Remove(subjectname.Text);
            }
            subject.Remove(Search.Text);
            var temp = new Thread(SubjectAnimate);
            temp.IsBackground = true;
            temp.Start();
        }

        private void MenuFreeDon_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\Вільні викладачі.xlsx"));
        }

        private void MenuFreeAud_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\Вільні аудиторії.xlsx"));
        }

        private void audsave_Click(object sender, RoutedEventArgs e)
        {
            audience.Remove(Search.SelectedItem.ToString());
            for (int x = 0; x < group.Count; x++)
                if (group.Values.ToArray()[x].relatedAud == Search.SelectedItem.ToString())
                    group[group.Keys.ToArray()[x]].relatedAud = audname.Text;
            for (int x = 0; x < subject.Count; x++)
                if (subject.Values.ToArray()[x].relatedAud == Search.SelectedItem.ToString())
                    subject[subject.Keys.ToArray()[x]].relatedAud = audname.Text;
            for (int x = 0; x < don.Count; x++)
                if (don.Values.ToArray()[x].relatedAud == Search.SelectedItem.ToString())
                    don[don.Keys.ToArray()[x]].relatedAud = audname.Text;
            audience.Remove(Search.SelectedItem.ToString());
            audience.Add(audname.Text);
            Search.SelectedIndex = -1;
            var temp = new Thread(AudAnimate);
            temp.IsBackground = true;
            temp.Start();
        }

        private void Day_RightClick(object sender, MouseButtonEventArgs e)
        {
            // проверить / в остальных, убрать, не вызывать анимэйшон если уже показаны чекбоксы
            CheckBox cbox = (CheckBox)sender;
            //MessageBox.Show((tbox == dond1).ToString());
            var cboxcs = cbox.Content.ToString();
            if (!cboxcs.Contains('/'))
            {
                if (ChosenDay != null)
                {
                    var cdcs = ChosenDay.Content.ToString();
                    ChosenDay.Content = cdcs.Substring(0, cdcs.Length - 1);
                }
                ChosenDay = cbox;
                cbox.Content = cboxcs + '/';
                var curday = don[donname.Text].dayStats[DaynameToNum(cboxcs)];
                donc1.IsChecked = curday[0];
                donc2.IsChecked = curday[1];
                donc3.IsChecked = curday[2];
                donc4.IsChecked = curday[3];
                donc5.IsChecked = curday[4];
                donc6.IsChecked = curday[5];
                if (!donc1.IsVisible)
                {
                    var temp1 = new Thread(CoupleExcAnimate);
                    temp1.IsBackground = true;
                    temp1.Start();
                }
            }
            else
            {
                ChosenDay = null;
                if (donc1.IsVisible)
                {
                    var temp1 = new Thread(CoupleExcAnimate);
                    temp1.IsBackground = true;
                    temp1.Start();
                }
                cbox.Content = cboxcs.Substring(0, cboxcs.Length - 1);
            }

            //switch (tbox.Content)
            //{
            //    case "Понеділок": ; break;
            //    case "Вівторок": break;
            //    case "Середа": break;
            //    case "Четвер": break;
            //    case "П'ятниця": break;
            //}
        }

        private void Couple_Click(object sender, RoutedEventArgs e)
        {
            var curdon = don[donname.Text];
            var chosendaynum = DaynameToNum(ChosenDay.Content.ToString());
            var curday = curdon.dayStats[chosendaynum];
            var cbox = (CheckBox)sender;
            curday[Convert.ToInt32(cbox.Uid)-1] = (bool)cbox.IsChecked;
            curdon.fixConsC();
            ChosenDay.IsChecked = curdon.possDays[chosendaynum];
        }

        private void dond_Click(object sender, RoutedEventArgs e)
        {
            CheckBox cbox = (CheckBox)sender;
            if (Convert.ToBoolean(cbox.IsChecked))
            {
                don[donname.Text].includeDay(cbox.Content.ToString());
            }
            else
            {
                don[donname.Text].excludeDay(cbox.Content.ToString());
            }
            if (donc1.IsVisible && cbox == ChosenDay)
            {
                var tbcs = cbox.Content.ToString();
                var curday = don[donname.Text].dayStats[DaynameToNum(tbcs)];
                donc1.IsChecked = curday[0];
                donc2.IsChecked = curday[1];
                donc3.IsChecked = curday[2];
                donc4.IsChecked = curday[3];
                donc5.IsChecked = curday[4];
                donc6.IsChecked = curday[5];
            }
        }
        #endregion

        #region Selection changed callbacks
        private void Search_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var sel = Search.SelectedItem;
            if (don.Keys.ToArray().Contains(sel))
            {
                var seldon = don[sel.ToString()];
                if (!donname.IsVisible)
                {
                    var temp = new Thread(DonAnimate);
                    temp.IsBackground = true;
                    temp.Start();
                }
                SortedSet<string> audinfo = audience;
                audinfo.Add("");
                donrelaud.ItemsSource = audinfo;
                donrelaud.SelectedItem = seldon.relatedAud;
                donrelatedsubjects.ItemsSource = seldon.relatedSubjects;
                donrelatedsubjects.Columns[0].Width = 320;
                donname.Text = sel.ToString();
                dond1.IsChecked = seldon.possDays[0];
                dond2.IsChecked = seldon.possDays[1];
                dond3.IsChecked = seldon.possDays[2];
                dond4.IsChecked = seldon.possDays[3];
                dond5.IsChecked = seldon.possDays[4];
            }
            else if (group.Keys.ToArray().Contains(sel))
            {
                if (!groupname.IsVisible)
                {
                    var temp = new Thread(GroupAnimate);
                    temp.IsBackground = true;
                    temp.Start();
                }
                var temp1 = group[sel.ToString()].relatedSubjects;
                List<List<string>> infoset = new List<List<string>>();
                for (int x = 0; x < temp1.Count; x++)
                {
                    infoset.Add(new List<string>());
                    infoset[x].AddRange(new List<string>() { "", "", "", "" });
                    infoset[x][0] = temp1.Keys.ToArray()[x];
                    for (int y = 1; y <= temp1.Values.ToArray()[x].Count; y++)
                        infoset[x][y] = temp1.Values.ToArray()[x][y - 1];
                }
                grouprelatedinfo.ItemsSource = infoset;
                groupname.Text = sel.ToString();
                grouprelaud.Text = group[sel.ToString()].relatedAud;
                groupstudyingweeks.Text = group[sel.ToString()].StudyingWeeks.ToString();
            }
            else if (subject.Keys.ToArray().Contains(sel))
            {
                if (!subjectname.IsVisible)
                {
                    var temp = new Thread(SubjectAnimate);
                    temp.IsBackground = true;
                    temp.Start();
                }
                subjectname.Text = subject[sel.ToString()].subjectName;
                subjectrelaud.Text = subject[sel.ToString()].relatedAud;
            }
            else if (audience.Contains(sel))
            {
                if (!audname.IsVisible)
                {
                    var temp = new Thread(AudAnimate);
                    temp.IsBackground = true;
                    temp.Start();
                }
                audname.Text = sel.ToString();
            }
        }

        private void grouprelatedsubjects_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (grouprelatedinfo.SelectedIndex != -1)
            {
                if (!grouprelsubjdel.IsVisible)
                {
                    var temp = new Thread(GroupRelSubjDelAnimate);
                    temp.IsBackground = true;
                    temp.Start();
                }
                if (!editsubjname.IsVisible)
                {
                    var temp = new Thread(GroupNewRelSubjAnimate);
                    temp.IsBackground = true;
                    temp.Start();
                }
                var temp1 = group[Search.SelectedItem.ToString()].relatedSubjects;
                editsubjname.Text = temp1.Keys.ToArray()[grouprelatedinfo.SelectedIndex];
                editdonname.Text = temp1.Values.ToArray()[grouprelatedinfo.SelectedIndex][0];
                editcoupleahalf.Text = temp1.Values.ToArray()[grouprelatedinfo.SelectedIndex][1];
                if (temp1.Values.ToArray()[grouprelatedinfo.SelectedIndex].Count == 3)
                    editseconddonname.Text = temp1.Values.ToArray()[grouprelatedinfo.SelectedIndex][2];
            }
        }

        private void donrelatedsubjects_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tmp = Mouse.GetPosition(System.Windows.Application.Current.MainWindow);
            if (!donrelsubjdel.IsVisible && tmp.X >= 213 && tmp.X < 741 && tmp.Y >= 219 && tmp.Y < 393)
            {
                var temp = new Thread(DonRelSubjDelAnimate);
                temp.IsBackground = true;
                temp.Start();
            }
        }

        private void donrelaud_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Search.SelectedIndex != -1)
                don[Search.SelectedItem.ToString()].relatedAud = donrelaud.SelectedItem.ToString();
        }
        #endregion

        #region Database update callbacks
        private void groupaddcouple_dbudon(object sender, DependencyPropertyChangedEventArgs e)
        {
            editdonname.ItemsSource = don.Keys.ToArray();
            editseconddonname.ItemsSource = don.Keys.ToArray();
        }

        private void groupaddcouple_dbusubj(object sender, DependencyPropertyChangedEventArgs e)
        {
            editsubjname.ItemsSource = subject.Keys.ToArray();
        }

        private void donrelsubjname_dbudrs(object sender, DependencyPropertyChangedEventArgs e)
        {
            donrelsubjname.ItemsSource = subject.Keys.ToArray();
        }
        #endregion
        #endregion

    }
}