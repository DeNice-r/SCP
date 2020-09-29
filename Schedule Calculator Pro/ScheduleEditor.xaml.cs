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

            List<StackPanel> uigroups = new List<StackPanel>();

            while (!Program.loadedinfo) Thread.Sleep(100);

            for(int _group = 0; _group < sched.Count(); _group++)
            {
                uigroups.Add(new StackPanel());
                for(int _day = 0; _day < 5; _day++)
                {
                    var ttblock = new TextBlock(); // temp. textblock
                    var ttblock1 = new TextBlock();
                    ttblock1.Text = Program.group.Keys.ToArray()[_group];
                    ttblock.Text = DOW[_day];
                    ttblock.TextAlignment = TextAlignment.Center;
                    var border1 = new Border();
                    var border2 = new Border();
                    border1.BorderBrush = Brushes.Black;
                    border1.BorderThickness = new Thickness(1);
                    border1.Child = ttblock;
                    border2.BorderBrush = Brushes.Black;
                    border2.BorderThickness = new Thickness(2);
                    border2.Child = ttblock1;
                    uigroups[_group].Children.Add(border1);
                    uigroups[_group].Children.Add(border2);
                    for (int _couple = 0; _couple < 6; _couple++)
                    {
                        var cttblock = new TextBlock(); // couple temp. textblock
                        var border = new Border();
                        border.BorderBrush = Brushes.Black;
                        border.BorderThickness = new Thickness(1);
                        border.Child = cttblock;
                        string s = "";
                        if (sched[_group][_day].Count() > _couple)
                        {
                            foreach (var field in sched[_group][_day][_couple])
                            {
                                s += field + " ";
                            }
                        }
                        cttblock.Text = s;
                        //cttblock.TextAlignment = TextAlignment.Center;
                        
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
                parent.Close();
            }

        }

        //private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        //{

        //}

        private void Window_SizeChanged(object sender, EventArgs e)
        {
            ContentMgr.Height = ActualHeight - 39;
            ContentMgr.Width = ActualWidth - 16;
        }
    }
}
