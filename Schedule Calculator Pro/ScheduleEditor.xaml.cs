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

            
            for(int _group = 0; _group < sched.Count(); _group++)
            {

            }
        }

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
