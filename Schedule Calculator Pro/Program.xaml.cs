//using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Animation;
using _Excel = Microsoft.Office.Interop.Excel;

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

        public Program()
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
                t++;
                while (excelTemp.BReadCell(t, 2))
                {
                    var tmp1 = excelTemp.ReadCell(t, 2);       // предмет
                    if (!subject.ContainsKey(tmp1))
                    {
                        subject.Add(tmp1, new Subject(tmp1));
                    }
                    if (!don[tmp].relatedSubjects.Contains(tmp1))
                        don[tmp].relatedSubjects.Add(tmp1);
                    if (excelTemp.BReadCell(t, 3))
                    {
                        subject[tmp1].relatedAud = excelTemp.ReadCell(t, 3);        // аудитория предмета
                        audience.Add(excelTemp.ReadCell(t, 3));
                    }
                    t++;
                    while (excelTemp.BReadCell(t, 4))
                    {
                        var tmp2 = excelTemp.ReadCell(t, 4);    // группа
                        if (!group.ContainsKey(tmp2))
                        {
                            group.Add(tmp2, new Group(tmp2));
                        }
                        if (!group[tmp2].relatedSubjects.ContainsKey(tmp1))
                        {
                            var tmp3 = "";
                            if (excelTemp.BReadCell(t, 5))
                                tmp3 = excelTemp.ReadCell(t, 5);            // кол-во пар в семестре
                            group[tmp2].SubjValEdit(tmp1, tmp, tmp3);
                        }
                        else if (!group[tmp2].relatedSubjects[tmp1].Contains(tmp))
                            group[tmp2].relatedSubjects[tmp1].Add(tmp);
                        if (excelTemp.BReadCell(t, 6))
                        {
                            group[tmp2].relatedAud = excelTemp.ReadCell(t, 6);
                            audience.Add(excelTemp.ReadCell(t, 6));
                        }
                        t++;
                    }
                }
            }

            var i = 0;

            while (excelTemp.BReadCell(i, 7))
            {
                var tmp = excelTemp.ReadCell(i, 7);
                if (!group.ContainsKey(tmp))
                {
                    group.Add(tmp, new Group(tmp));
                }
                i++;
            }
            i = 0;
            while (excelTemp.BReadCell(i, 8))
            {
                var tmp = excelTemp.ReadCell(i, 8);
                if (!subject.ContainsKey(tmp))
                {
                    subject.Add(tmp, new Subject(tmp));
                }
                i++;
            }
            i = 0;
            while (excelTemp.BReadCell(i, 9))
            {
                var tmp = excelTemp.ReadCell(i, 9);
                if (!group.ContainsKey(tmp))
                {
                    audience.Add(tmp);
                }
                i++;
            }

            while (excelTemp.BReadCell(i, 10))
            {
                var tmp = excelTemp.ReadCell(i, 10);

                {
                    group[tmp].StudyingWeeks = Convert.ToInt32(excelTemp.ReadCell(i, 11));
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
                excelTemp1.WriteToCell(t, 0, don.Keys.ToArray()[j]);
                excelTemp1.WriteToCell(t, 1, don.Values.ToArray()[j].relatedAud);

                if (unwritten[2].Contains(don.Values.ToArray()[j].relatedAud))
                    unwritten[2].Remove(don.Values.ToArray()[j].relatedAud);

                t++;
                for (int k = 0; k < don.Values.ToArray()[j].relatedSubjects.Count; k++)
                {
                    excelTemp1.WriteToCell(t, 2, don.Values.ToArray()[j].relatedSubjects[k]);
                    if (unwritten[1].Contains(don.Values.ToArray()[j].relatedSubjects[k]))
                        unwritten[1].Remove(don.Values.ToArray()[j].relatedSubjects[k]);

                    if (!subject.ContainsKey(don.Values.ToArray()[j].relatedSubjects[k]))
                        MessageBox.Show(don.Values.ToArray()[j].relatedSubjects[k]);

                    excelTemp1.WriteToCell(t, 3, subject[don.Values.ToArray()[j].relatedSubjects[k]].relatedAud);

                    if (unwritten[2].Contains(subject[don.Values.ToArray()[j].relatedSubjects[k]].relatedAud))
                        unwritten[2].Remove(subject[don.Values.ToArray()[j].relatedSubjects[k]].relatedAud);
                    t++;
                    for (int g = 0; g < group.Count; g++)
                    {
                        if (group.Values.ToArray()[g].relatedSubjects.ContainsKey(don.Values.ToArray()[j].relatedSubjects[k]) && group.Values.ToArray()[g].relatedSubjects.Values.ToArray().Any(x => x.Contains(don.Keys.ToArray()[j])))
                        {
                            excelTemp1.WriteToCell(t, 4, group.Keys.ToArray()[g]);
                            if (unwritten[0].Contains(group.Keys.ToArray()[g]))
                                unwritten[0].Remove(group.Keys.ToArray()[g]);

                            excelTemp1.WriteToCell(t, 5, group.Values.ToArray()[g].relatedSubjects[don.Values.ToArray()[j].relatedSubjects[k]][1].ToString());
                            excelTemp1.WriteToCell(t, 6, group.Values.ToArray()[g].relatedAud);

                            if (unwritten[2].Contains(group.Values.ToArray()[g].relatedAud))
                                unwritten[2].Remove(group.Values.ToArray()[g].relatedAud);

                            t++;
                        }
                    }
                }
            }

            for (int i = 0; i < unwritten[0].Count; i++)
            {
                excelTemp1.WriteToCell(i, 7, unwritten[0][i]);
            }
            for (int i = 0; i < unwritten[1].Count; i++)
            {
                excelTemp1.WriteToCell(i, 8, unwritten[1][i]);
            }
            for (int i = 0; i < unwritten[2].Count; i++)
            {
                excelTemp1.WriteToCell(i, 9, unwritten[2][i]);
            }

            var temp1 = group.Keys.ToArray();
            var temp2 = group.Values.ToArray();
            for (int i = 0; i < temp1.Count(); i++)
            {
                excelTemp1.WriteToCell(i, 10, temp1[i]);
                excelTemp1.WriteToCell(i, 11, temp2[i].StudyingWeeks.ToString());
            }

            excelTemp1.ws.Columns.AutoFit();
            excelTemp1.SaveAs();
            excelTemp1.close();
            MessageBox.Show("Дані збережено.");
        }

        private void SearchHandle(object sender, EventArgs e)
        {
            if (!(don.Keys.ToList().All(x => database.Contains(x)) && group.Keys.ToList().All(x => database.Contains(x)) && subject.Keys.ToList().All(x => database.Contains(x)) && audience.All(x => database.Contains(x))))
            {
                database.RemoveRange(0, database.Count);
                database.AddRange(don.Keys.ToArray());
                database.AddRange(group.Keys.ToArray());
                database.AddRange(subject.Keys.ToArray());
                database.AddRange(audience.ToArray());
            }
            Search.ItemsSource = database.Where(x => x.ToLower().Contains(Search.Text.ToLower())).ToArray();
            if (!Search.IsDropDownOpen)
                Search.IsDropDownOpen = true;
        }


        //       Animations
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

        private void DonAnimate()
        {
            if (donname.IsVisible)
            {
                //DoubleAnimation anim1 = new DoubleAnimation(0, TimeSpan.FromMilliseconds(500));
                //donname.BeginAnimation(TextBox.OpacityProperty, anim1);
                //donrelatedsubjects.BeginAnimation(TextBox.OpacityProperty, anim1);
                //deletedon.BeginAnimation(TextBox.OpacityProperty, anim1);
                //newdonrelsubj.BeginAnimation(TextBox.OpacityProperty, anim1);
                for (double x = .98; x >= 0; x -= .02)
                {
                    donname.Dispatcher.Invoke(delegate { donname.Opacity = x; });
                    donrelatedsubjects.Dispatcher.Invoke(delegate { donrelatedsubjects.Opacity = x; });
                    deletedon.Dispatcher.Invoke(delegate { deletedon.Opacity = x; });
                    newdonrelsubj.Dispatcher.Invoke(delegate { newdonrelsubj.Opacity = x; });
                    Thread.Sleep(2);
                }
                donname.Dispatcher.Invoke(delegate { donname.Visibility = Visibility.Hidden; donname.Text = ""; });
                donrelatedsubjects.Dispatcher.Invoke(delegate { donrelatedsubjects.Visibility = Visibility.Hidden; donrelatedsubjects.ItemsSource = null; });
                deletedon.Dispatcher.Invoke(delegate { deletedon.Visibility = Visibility.Hidden; });
                newdonrelsubj.Dispatcher.Invoke(delegate { newdonrelsubj.Visibility = Visibility.Hidden; });
            }
            else
            {
                donname.Dispatcher.Invoke(delegate { donname.Visibility = Visibility.Visible; });
                donrelatedsubjects.Dispatcher.Invoke(delegate { donrelatedsubjects.Visibility = Visibility.Visible; });
                deletedon.Dispatcher.Invoke(delegate { deletedon.Visibility = Visibility.Visible; });
                newdonrelsubj.Dispatcher.Invoke(delegate { newdonrelsubj.Visibility = Visibility.Visible; });

                //DoubleAnimation anim1 = new DoubleAnimation(1, TimeSpan.FromMilliseconds(500));
                //donname.BeginAnimation(TextBox.OpacityProperty, anim1);
                //donrelatedsubjects.BeginAnimation(TextBox.OpacityProperty, anim1);
                //deletedon.BeginAnimation(TextBox.OpacityProperty, anim1);
                //newdonrelsubj.BeginAnimation(TextBox.OpacityProperty, anim1);
                for (double x = .02; x < 1; x += .02)
                {
                    donname.Dispatcher.Invoke(delegate { donname.Opacity = x; });
                    donrelatedsubjects.Dispatcher.Invoke(delegate { donrelatedsubjects.Opacity = x; });
                    deletedon.Dispatcher.Invoke(delegate { deletedon.Opacity = x; });
                    newdonrelsubj.Dispatcher.Invoke(delegate { newdonrelsubj.Opacity = x; });
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

        }



        //          Callbacks

        // Click callbacks
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


        }

        private void save_Click(object sender, RoutedEventArgs e)
        {
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
            System.Diagnostics.Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\SCP\\Розклад 1 курс.xlsx");
            System.Diagnostics.Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\SCP\\Розклад 2 курс.xlsx");
            System.Diagnostics.Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\SCP\\Розклад 3 курс.xlsx");
            System.Diagnostics.Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\SCP\\Розклад 4 курс.xlsx");
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
                if(group[Search.SelectedItem.ToString()].relatedSubjects.Keys.ToArray()[grouprelatedinfo.SelectedIndex] != editsubjname.Text)
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
            System.Diagnostics.Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\SCP\\Вільні викладачі.xlsx");
        }

        private void MenuFreeAud_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\SCP\\Вільні аудиторії.xlsx");
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

        private void Day_DoubleClick(object sender, MouseButtonEventArgs e)
        {
            CheckBox tbox = (CheckBox)sender;
            var tbcs = tbox.Content.ToString();
            if (!tbcs.Contains('/')) {
                tbox.Content = tbcs + '/';
            }
            else {
                tbox.Content = tbcs.Substring(0, tbcs.Length - 1);
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

        }


        // Selection changed callbacks
        private void Search_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (don.Keys.ToArray().Contains(Search.SelectedItem))
            {
                if (!donname.IsVisible)
                {
                    var temp = new Thread(DonAnimate);
                    temp.IsBackground = true;
                    temp.Start();
                }
                donrelatedsubjects.ItemsSource= don[Search.SelectedItem.ToString()].relatedSubjects;
                donrelatedsubjects.Columns[0].Width = 320;
                donname.Text = Search.SelectedItem.ToString();
            }
            else if (group.Keys.ToArray().Contains(Search.SelectedItem))
            {
                if (!groupname.IsVisible)
                {
                    var temp = new Thread(GroupAnimate);
                    temp.IsBackground = true;
                    temp.Start();
                }
                var temp1 = group[Search.SelectedItem.ToString()].relatedSubjects;
                List<List<string>> infoset = new List<List<string>>();
                for (int x = 0; x < temp1.Count; x++)
                {
                    infoset.Add(new List<string>());
                    infoset[x].AddRange(new List<string>() { "", "", "", "" });
                    infoset[x][0] = temp1.Keys.ToArray()[x];
                    for (int y = 1; y <= temp1.Values.ToArray()[x].Count; y++)
                        infoset[x][y] = temp1.Values.ToArray()[x][y-1];
                }
                grouprelatedinfo.ItemsSource = infoset;
                groupname.Text = Search.SelectedItem.ToString();
                grouprelaud.Text = group[Search.SelectedItem.ToString()].relatedAud;
                groupstudyingweeks.Text = group[Search.SelectedItem.ToString()].StudyingWeeks.ToString();
            }
            else if (subject.Keys.ToArray().Contains(Search.SelectedItem))
            {
                if (!subjectname.IsVisible)
                {
                    var temp = new Thread(SubjectAnimate);
                    temp.IsBackground = true;
                    temp.Start();
                }
                subjectname.Text = subject[Search.SelectedItem.ToString()].subjectName;
                subjectrelaud.Text = subject[Search.SelectedItem.ToString()].relatedAud;
            }
            else if(audience.Contains(Search.SelectedItem))
            {
                if (!audname.IsVisible)
                {
                    var temp = new Thread(AudAnimate);
                    temp.IsBackground = true;
                    temp.Start();
                }
                audname.Text = Search.SelectedItem.ToString();
            }
        }

        private void grouprelatedsubjects_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
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

        private void donrelatedsubjects_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var tmp = Mouse.GetPosition(System.Windows.Application.Current.MainWindow);
            if (!donrelsubjdel.IsVisible && tmp.X >= 213 && tmp.X < 741 && tmp.Y >= 219 && tmp.Y < 393) {
                var temp = new Thread(DonRelSubjDelAnimate);
                temp.IsBackground = true;
                temp.Start();
            }
        }


        // ?
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

        ~Program()
        {
            Excel.Kill("База даних.xlsx");
            Excel.Kill("Дані.xlsx");
        }

    }
}