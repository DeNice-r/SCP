using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Input;
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

        public static void SettingsHandle()
        {
            //Перевіряємо, чи існує файл з налаштуваннями, відкриваємо його, якщо він є і створюємо у зворотньому випадку.
            if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\SCP\\Settings.xlsx"))
                GetSettings();
            else if(File.Exists("RawData.xlsx")){
                CreateSettings();
            }
            else {
                MessageBox.Show("Створіть файл з початковими даними і назвіть його RawData.xlsx");
            }
        }

        private static void GetSettings()
        {
            Excel excelTemp = new Excel(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\SCP\\Settings.xlsx", 1);
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
            MessageBox.Show("Завантаження початкових даних завершено.");
        }

        private static void CreateSettings()
        {
            if (!File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\SCP\\Settings.xlsx"))
                {
                    Excel excelTemp = new Excel(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\RawData.xlsx"), 1);

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
            Excel excelTemp1 = new Excel(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\SCP\\Settings.xlsx");
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
            MessageBox.Show("Saving done.");
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
            if(!Search.IsDropDownOpen)
                Search.IsDropDownOpen = true;
        }

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

        private void donrelatedsubjects_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var tmp = Mouse.GetPosition(System.Windows.Application.Current.MainWindow);
            if (!donrelsubjdel.IsVisible && tmp.X >= 213 && tmp.X < 741 && tmp.Y >= 219 && tmp.Y < 393) {
                var temp = new Thread(DonRelSubjDelAnimate);
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
    }

    public class Schedule
    {
        public List<List<List<List<string>>>> schedule { get; set; } = new List<List<List<List<string>>>>();        // Группа > День > Пара > Преподаватели/Предметы/Аудитории
        public List<List<List<List<string>>>> scheduleFree { get; set; } = new List<List<List<List<string>>>>();        // День > Пара >  Преподаватель/Аудитория
        private Random rnd = new Random();
        public Schedule() { }

        public void Start()
        {
            PreGen();
            for (int don = 0; don < Program.don.Count; don++)
            {
                var tgroups = new List<int>();
                var tlist = new List<List<int>>();
                for (int group = 0; group < Program.group.Count; group++)
                {
                    var tgroup = Program.group.Values.ToArray()[group];
                    for (int x = 0; x < tgroup.relatedSubjects.Count; x++)
                    {
                        if (tgroup.relatedSubjects.Values.ToArray()[x][0] == Program.don.Values.ToArray()[don].donName)
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
                        if (scheduleFree[day][couple][0].Contains(Program.don.Keys.ToArray()[don]))
                            for (int group = 0; group < tgroups.Count; group++)
                            {
                                if (Program.group.Values.ToArray()[tgroups[group]].couplesXdayGet(day) <= couple || schedule[tgroups[group]][day][couple].Count != 0)
                                    continue;
                                for (int subj = 0; subj < tlist[group].Count; subj++)
                                    if (Convert.ToInt32(Program.group.Values.ToArray()[tgroups[group]].relatedSubjects.Values.ToArray()[tlist[group][subj]][1]) > 0)
                                    {
                                        schedule[tgroups[group]][day][couple].Add(Program.group.Values.ToArray()[tgroups[group]].relatedSubjects.Keys.ToArray()[tlist[group][subj]]); // Предмет
                                        schedule[tgroups[group]][day][couple].Add(Program.don.Keys.ToArray()[don]); // Преподаватель
                                        var b = Program.group.Values.ToArray()[tgroups[group]].relatedSubjects.Values.ToArray()[tlist[group][subj]].Count == 3 && Program.group.Values.ToArray()[tgroups[group]].relatedSubjects.Values.ToArray()[tlist[group][subj]][2] != "";
                                        if (b)
                                        {
                                            schedule[tgroups[group]][day][couple].Add(Program.group.Values.ToArray()[tgroups[group]].relatedSubjects.Values.ToArray()[tlist[group][subj]][2]);
                                            scheduleFree[day][couple][0].Remove(Program.group.Values.ToArray()[tgroups[group]].relatedSubjects.Values.ToArray()[tlist[group][subj]][2]);
                                        }
                                        schedule[tgroups[group]][day][couple].Add(getAud(day, couple, Program.group.Values.ToArray()[tgroups[group]], Program.group.Values.ToArray()[tgroups[group]].relatedSubjects.Keys.ToArray()[tlist[group][subj]], 0)); // Аудитория
                                        if (b)
                                            schedule[tgroups[group]][day][couple].Add(getAud(day, couple, Program.group.Values.ToArray()[tgroups[group]], Program.group.Values.ToArray()[tgroups[group]].relatedSubjects.Keys.ToArray()[tlist[group][subj]], 0)); // Аудитория
                                        scheduleFree[day][couple][0].Remove(Program.don.Keys.ToArray()[don]);
                                        Program.group.Values.ToArray()[tgroups[group]].relatedSubjects.Values.ToArray()[tlist[group][subj]][1] = (Convert.ToInt32(Program.group.Values.ToArray()[tgroups[group]].relatedSubjects.Values.ToArray()[tlist[group][subj]][1]) - 16).ToString();
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
            for (int _day = 0; _day < 5; _day++)
            {
                scheduleFree.Add(new List<List<List<string>>>());
                for (int _couple = 0; _couple < 6; _couple++)
                {

                    scheduleFree[_day].Add(new List<List<string>>());
                    scheduleFree[_day][_couple].Add(new List<string>());
                    scheduleFree[_day][_couple][0] = Program.don.Keys.ToList();
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

        private void Save()
        {
            for (int _course = 0; _course < 4; _course++)
            {

                var xcl = new Excel(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\SCP\\Розклад " + (_course + 1) + " курс.xlsx");
                var col = 2;
                var row = 1;
                var mrow = 0;
                var mcol = 0;
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

                for(int crs = 0; crs < 4; crs++)
                {
                    wtf.Add(new List<List<int>>());
                    for(int dy = 0; dy < 5; dy++)
                    {
                        wtf[crs].Add(new List<int>());
                        for(int cpl = 0; cpl < 6; cpl++)
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
                            for(int grp = 0; grp < Program.group.Count; grp++)
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
                    xcl.ws.Range[xcl.ws.Cells[t1, 1], xcl.ws.Cells[t1 + wtf[_course][zz].Sum()-1, 1]].Merge();
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
            Excel xcl1 = new Excel(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\SCP\\Вільні викладачі.xlsx");
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
                        xrow += cxz[tx]+2;
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
                    xcl1.ws.Range[xcl1.ws.Cells[xrowt+1, xcol+1], xcl1.ws.Cells[xrow, xcol+1]].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range[xcl1.ws.Cells[xrowt+1, xcol+1], xcl1.ws.Cells[xrow, xcol+1]].Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range[xcl1.ws.Cells[xrowt+1, xcol+1], xcl1.ws.Cells[xrow, xcol+1]].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                    xcl1.ws.Range[xcl1.ws.Cells[xrowt+1, xcol+1], xcl1.ws.Cells[xrow, xcol+1]].Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
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

            xcl1 = new Excel(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\SCP\\Вільні аудиторії.xlsx");
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
            while(inp > 0)
            {
                outp = Convert.ToChar((inp - 1) % 26 + 'A') + outp;
                inp = (inp - 1) / 26;
            }
            return outp;
        }

        private int func1(List<int> inp, int crs, int dy)
        {
            var t = 5;
            while (inp[t] == 0)
                t--;
            return t;
        }

        private int func2(List<List<int>> inp)
        {
            var outp = 0;

            for(int x = 0; x < inp.Count; x++)
            {
                for(int y = 0; y < inp[x].Count; y++)
                {
                    outp += inp[x][y];
                }
            }
            return outp;
        }
    }

    public class Excel
    {
        string path = "";
        public _Application excel = new _Excel.Application();
        public Workbook wb;
        public Worksheet ws;
        public Excel(string path)
        {
            Kill(path);
            this.path = path;
            excel.SheetsInNewWorkbook = 1;
            wb = excel.Workbooks.Add(1);
            ws = wb.Worksheets[1];
        }
        public Excel(string path, int sheet)
        {
            Kill(path);
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }
        public void close()
        {
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            Kill(path);
        }

        public string ReadCell(int i, int j)
        {
            return Convert.ToString(ws.Cells[i + 1, j + 1].Value2);
        }
        public void WriteToCell(int i, int j, string s)
        {
            ws.Cells[i + 1, j + 1].Value2 = s;
        }
        public void Save()
        {
            wb.Save();
        }
        public void SaveAs()
        {
            if (File.Exists(path))
                File.Delete(path);
            wb.SaveAs(path);
        }
        public bool BReadCell(int i, int j)
        {
            i++; j++;
            if (ws.Cells[i, j].Value2 != null)
                return true;
            else
                return false;
        }

        private static void Kill(string excelFileName) // убиваем процес по имени файла
        {
            var processes = from p in Process.GetProcessesByName("EXCEL") select p;

            foreach (var process in processes)
                if (process.MainWindowTitle == "Microsoft Excel - " + excelFileName)
                    process.Kill();
        }
    }

    public class Don
    {
        public string donName { get; set; }
        public List<string> relatedSubjects { get; set; } = new List<string>();
        public string relatedAud { get; set; } = "";
        public Don(string donName)
        {
            this.donName = donName;
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