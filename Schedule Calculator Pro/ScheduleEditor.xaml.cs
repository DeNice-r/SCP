using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.IO;
using System.Threading;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Schedule_Calculator_Pro
{
    /// <summary>
    /// Логика взаимодействия для ScheduleEditor.xaml
    /// </summary>
    public partial class ScheduleEditor : Window
    {
        private static Program parent = null;
        public static string[] DOW = { "Понеділок", "Вівторок", "Середа", "Четвер", "П'ятниця" };
        List<StackPanel> uigroups = new List<StackPanel>();

        public ScheduleEditor(Program p)
        {
            InitializeComponent();
            parent = p;
            ContentMgr.Height = Height - 39;
            ContentMgr.Width = Width - 16;

            var sched = Program.schedule.schedule;

            // Table generation
            if (File.Exists(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\schedcfg.json")))
            {
                sched = (List<List<List<List<string>>>>)JsonSerializer.Deserialize(System.IO.File.ReadAllText(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\schedcfg.json")), sched.GetType());
            }
            // Группа > День > Пара > Преподаватели/Предметы/Аудитории

            

            while (!Program.loadedinfo) Thread.Sleep(100);

            for(int _group = 0; _group < sched.Count(); _group++)
            {
                int lensubj = 0;
                int lendon = 0;
                int lenaud = 0;
                foreach (var d in sched[_group])
                    foreach (var c in d)
                    {
                        if (c.Count() > 0)
                        {
                            if (lensubj < c[0].Length)
                                lensubj = c[0].Length;
                            if (lendon < c[1].Length)
                                lendon = c[1].Length;

                            if (c.Count() > 3)
                            {
                                if (lendon < c[2].Length)
                                    lendon = c[2].Length;
                                if (lenaud < c[3].Length)
                                    lenaud = c[3].Length;
                            }
                            else if(lenaud < c[2].Length)
                                lenaud = c[2].Length;

                        }
                    }
                lensubj = lensubj * 7;
                lendon= lendon * 7;
                lenaud = lenaud * 7 + 10;
                var border2 = new Border();
                border2.BorderBrush = Brushes.Black;
                border2.BorderThickness = new Thickness(1);
                var label1 = new Label();
                border2.Child = label1 ;
                uigroups.Add(new StackPanel());
                label1.Content = Program.group.Keys.ToArray()[_group];
                label1.HorizontalContentAlignment = HorizontalAlignment.Center;
                label1.FontSize = 20;
                uigroups[_group].Children.Add(border2);
                for (int _day = 0; _day < 5; _day++)
                {
                    var label = new Label(); // temp. textblock
                    label.Content = DOW[_day];
                    label.HorizontalContentAlignment = HorizontalAlignment.Center;
                    var border1 = new Border();
                    border1.BorderBrush = Brushes.Black;
                    border1.BorderThickness = new Thickness(1);
                    //border1.Child = label;
                    uigroups[_group].Children.Add(border1);
                    for (int _couple = 0; _couple < 6; _couple++)
                    {
                        DockPanel dp = new DockPanel();
                        var border = new Border();
                        border.BorderBrush = Brushes.Black;
                        border.BorderThickness = new Thickness(1);
                        border.Child = dp;
                        if (sched[_group][_day].Count() > _couple)
                        {
                            var curcpl = sched[_group][_day][_couple];
                            var curlen = sched[_group][_day][_couple].Count();
                            if (curlen != 0)
                            {

                                var tbsubj = new Label();
                                var sp1 = new StackPanel(); var bord1 = new Border(); bord1.BorderBrush = Brushes.Black; bord1.BorderThickness = new Thickness(1, 0, 1, 0);
                                var sp2 = new StackPanel(); var bord2 = new Border(); bord2.BorderBrush = Brushes.Black; bord2.BorderThickness = new Thickness(1, 0, 1, 0);
                                var tbdon1 = new Label(); bord1.Child = tbdon1; sp1.Children.Add(bord1);
                                var tbdon2 = new Label(); bord2.Child = tbdon2; sp1.Children.Add(bord2);
                                var tbaud1 = new Label(); sp2.Children.Add(tbaud1);
                                var tbaud2 = new Label(); sp2.Children.Add(tbaud2);
                                tbsubj.Width = lensubj; tbsubj.Height = 35; tbsubj.VerticalContentAlignment = VerticalAlignment.Center;
                                tbdon1.Width = tbdon2.Width = lendon;
                                tbaud1.Width = tbaud2.Width = lenaud;

                                tbsubj.Content = curcpl[0];
                                tbdon1.Content = curcpl[1];
                                if (curlen > 3)
                                {
                                    tbdon2.Content = curcpl[2];
                                    tbaud1.Content = curcpl[3];
                                    tbaud2.Content = curcpl[4];
                                    tbdon1.Height = 23;
                                    tbaud1.Height = 23;
                                }
                                else
                                {
                                    tbdon1.Height = 50; tbdon1.VerticalContentAlignment = VerticalAlignment.Center;
                                    tbaud1.Height = 50; tbaud1.VerticalContentAlignment = VerticalAlignment.Center;
                                    tbaud1.Content = curcpl[2];
                                }

                                dp.Children.Add(tbsubj);
                                dp.Children.Add(sp1);
                                dp.Children.Add(sp2);
                            }
                        }
                        dp.MouseLeftButtonDown += Cpl_DragEnter;
                        dp.Height = 50;
                        uigroups[_group].Children.Add(border);
                    }
                }
            }
            foreach(var group in uigroups)
                docks.Children.Add(group);
        } // Перетягивание на коллбеке дрега с переводом сендера в текстблок.

        private void Window_Closed(object sender, EventArgs e)
        {
            Program.scheditwin = null;
            if (Program.workwithschedit)
            {
                if(parent != null)
                    parent.Close();
            }

        }

        private void Cpl_DragEnter(object sender, MouseButtonEventArgs e)
        {
            var dp = (DockPanel)sender;
            //MessageBox.Show("");
            //if()
        }

        private void Window_SizeChanged(object sender, EventArgs e)
        {
            ContentMgr.Height = ActualHeight - 39;
            ContentMgr.Width = ActualWidth - 16;
            //if(ContentMgr.Width > 1000 )
            //    MessageBox.Show(((TextBlock)((Border)uigroups[0].Children[0]).Child).ActualWidth.ToString());
        }

        private void segrid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.R)
            {
                segrid.UpdateLayout();
            }
        }
    }
}
