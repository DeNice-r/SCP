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
using drawing = System.Drawing;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Drawing.Imaging;
using System.Windows.Media.Effects;
using System.Xml.Schema;
using System.Runtime.CompilerServices;

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
        public static List<List<List<List<string>>>> sched = new List<List<List<List<string>>>>();

        public ScheduleEditor(Program p)
        {
            InitializeComponent();
            parent = p;
            ContentMgr.Height = Height - 39;
            ContentMgr.Width = Width - 16;
            // Table generation
            if (File.Exists(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\schedcfg.json")))
            {
                sched = (List<List<List<List<string>>>>)JsonSerializer.Deserialize(System.IO.File.ReadAllText(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\schedcfg.json")), sched.GetType());
                Program.schedule.scheduleFree = (List<List<List<List<string>>>>)JsonSerializer.Deserialize(System.IO.File.ReadAllText(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\schedfreecfg.json")), Program.schedule.scheduleFree.GetType());
                //Program.subject = (SortedDictionary<string, Subject>)JsonSerializer.Deserialize(System.IO.File.ReadAllText(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\schedsubj.json")), Program.subject.GetType());
                //Program.group = (SortedDictionary<string, Group>)JsonSerializer.Deserialize(System.IO.File.ReadAllText(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\schedgroup.json")), Program.group.GetType());
                //Program.don = (SortedDictionary<string, Don>)JsonSerializer.Deserialize(System.IO.File.ReadAllText(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\scheddon.json")), Program.don.GetType());
                Program.audience = (SortedSet<string>)JsonSerializer.Deserialize(System.IO.File.ReadAllText(System.Reflection.Assembly.GetExecutingAssembly().Location.Replace("\\Schedule Calculator Pro.exe", "\\schedaud.json")), Program.audience.GetType());
                Schedentry.sched = Program.schedule.schedule = sched;
            }
            // Группа > День > Пара > Преподаватели/Предметы/Аудитории



            while (!Program.loadedinfo) Thread.Sleep(100);

            for (int _group = 0; _group < sched.Count(); _group++)
            {
                {
                    int lensubj = 0, lendon = 0, lenaud = 0;
                    foreach (var d in sched[_group]) // Calc 
                        foreach (var c in d)
                        {
                            if (c.Count() > 0)
                            {
                                if (c.Count() == 6)
                                {
                                    if (lensubj < c[0].Length) lensubj = c[0].Length;
                                    if (lensubj < c[1].Length) lensubj = c[1].Length;
                                    if (lendon < c[2].Length) lendon = c[2].Length;
                                    if (lendon < c[3].Length) lendon = c[3].Length;
                                    if (lenaud < c[4].Length) lenaud = c[4].Length;
                                    if (lenaud < c[5].Length) lenaud = c[5].Length;
                                }
                                else if (c.Count() == 5)
                                {
                                    if (lensubj < c[0].Length) lensubj = c[0].Length;
                                    if (lendon < c[1].Length) lendon = c[1].Length;
                                    if (lendon < c[2].Length) lendon = c[2].Length;
                                    if (lenaud < c[3].Length) lenaud = c[3].Length;
                                    if (lenaud < c[4].Length) lenaud = c[4].Length;
                                }
                                else if (c.Count() == 3)
                                {
                                    if (lensubj < c[0].Length) lensubj = c[0].Length;
                                    if (lendon < c[1].Length) lendon = c[1].Length;
                                    if (lenaud < c[2].Length) lenaud = c[2].Length;
                                }
                            }
                        }
                    lensubj *= 7; lendon *= 7; lenaud = lenaud * 7 + 10;
                    Schedentry.colwidth.Add(new List<int>() { lensubj, lendon, lenaud });
                }
                Schedentry groupentry = new Schedentry(new string[] { Program.group.Keys.ToArray()[_group] }, new int[] { _group });

                uigroups.Add(new StackPanel());
                uigroups[_group].Children.Add(groupentry);
                for (int _day = 0; _day < 5; _day++)
                {
                    Schedentry dayentry = new Schedentry(new string[] { DOW[_day] }, new int[] { _group, _day });
                    uigroups[_group].Children.Add(dayentry);
                    for (int _couple = 0; _couple < 6; _couple++)
                    {
                        Schedentry sentry;
                        if (sched[_group][_day].Count() > _couple)
                            sentry = new Schedentry(sched[_group][_day][_couple].ToArray(), new int[] { _group, _day, _couple });
                        else
                            sentry = new Schedentry(new string[] { }, new int[] { _group, _day, _couple });
                        uigroups[_group].Children.Add(sentry);
                    }
                }
            }
            foreach (var group in uigroups)
                docks.Children.Add(group);
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            Program.scheditwin = null;
            if (Program.workwithschedit)
            {
                if (parent != null)
                    parent.Close();
            }

        }

        private void Window_SizeChanged(object sender, EventArgs e)
        {
            ContentMgr.Height = ActualHeight - 39;
            ContentMgr.Width = ActualWidth - 46;
            sidebar.Height = ActualHeight - 39;
            //if(ContentMgr.Width > 1000 )
            //    MessageBox.Show(((TextBlock)((Border)uigroups[0].Children[0]).Child).ActualWidth.ToString());
        }

        private void segrid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.R) // TODO: Скринить контент мгр и скроллить, потом клеить или скринить док/стек панели и клеить. Начинать со второго!!!
            {
                segrid.UpdateLayout();

                int finwidth = 0, finheight = 0;
                List<List<RenderTargetBitmap>> rtblist = new List<List<RenderTargetBitmap>>();

                foreach (StackPanel d in docks.Children)
                {
                    ContentMgr.ScrollToVerticalOffset(0);
                    rtblist.Add(new List<RenderTargetBitmap>());
                    finwidth += (int)((Border)d.Children[0]).ActualWidth;
                    foreach (Border s in d.Children)
                    {
                        RenderTargetBitmap rtb = new RenderTargetBitmap((int)s.ActualWidth, (int)s.ActualHeight, 96, 96, PixelFormats.Pbgra32);
                        rtb.Render(s);
                        rtblist.Last().Add(rtb);
                    }
                }

                foreach (Border b in ((StackPanel)docks.Children[0]).Children)
                    finheight += (int)b.ActualHeight;

                drawing.Bitmap sched = new drawing.Bitmap(finwidth, finheight);

                using (drawing.Graphics g = drawing.Graphics.FromImage(sched))
                {
                    int curwidth = 0, curheight = 0;
                    for (int x = 0; x < rtblist.Count(); x++)
                    {
                        curwidth += (int)((StackPanel)docks.Children[x]).ActualWidth;
                        curheight = 0;
                        for (int y = 1; y < rtblist[x].Count(); y++)
                        {
                            var rtb = rtblist[x][y];
                            drawing.Bitmap bm = new drawing.Bitmap(GetBitmap(rtb));
                            g.DrawImage(bm, curwidth, curheight);
                            curheight += (int)((Border)((StackPanel)docks.Children[x]).Children[y]).ActualHeight;
                        }
                    }
                }

                using (Stream fileStream = File.Create("sched.png"))
                {
                    sched.Save(fileStream, ImageFormat.Png);
                }
            }
        }

        public drawing.Bitmap GetBitmap(BitmapSource source)
        {
            drawing.Bitmap bmp = new drawing.Bitmap
            (
              source.PixelWidth,
              source.PixelHeight,
              System.Drawing.Imaging.PixelFormat.Format32bppPArgb
            );

            BitmapData data = bmp.LockBits
            (
                new System.Drawing.Rectangle(System.Drawing.Point.Empty, bmp.Size),
                ImageLockMode.WriteOnly,
                System.Drawing.Imaging.PixelFormat.Format32bppPArgb
            );

            source.CopyPixels
            (
              Int32Rect.Empty,
              data.Scan0,
              data.Height * data.Stride,
              data.Stride
            );

            bmp.UnlockBits(data);

            return bmp;
        }

        private void desel_Click(object sender, RoutedEventArgs e)
        {
            deselect();
        }

        public static void deselect()
        {
            if (Schedentry.cut != null)
            {
                Schedentry.cut.Opacity = 1;
                Schedentry.cut = null;
            }
        }

        private void remove_Click(object sender, RoutedEventArgs e)
        {
            if (Schedentry.cut != null)
            {
                var cut = Schedentry.cut;
                var p = (StackPanel)cut.Parent;

                var index = p.Children.IndexOf(cut);
                p.Children.RemoveAt(index);
                p.Children.Insert(index, new Schedentry(new string[] { }, cut.idx));
                Schedentry.cut = null;
            }
        }

        private void schedsave_Click(object sender, RoutedEventArgs e)
        {
            deselect();
            var usched = new List<List<List<List<string>>>>();
            foreach (StackPanel sp in docks.Children)
            {
                usched.Add(new List<List<List<string>>>());
                foreach (Schedentry se in sp.Children)
                {
                    if (se.subj[0].Content != null && DOW.Contains(se.subj[0].Content.ToString()))
                        usched.Last().Add(new List<List<string>>());
                    if (se.info.Length >= 3 || se.info.Length == 0)
                        usched.Last().Last().Add(new List<string>(se.info));
                }
            }
            {
            //string s = "";
            //var sched = Program.schedule.schedule;

            //double ovall = 0, succ = 0;
            //for (int x = 0; x < usched.Count(); x++)
            //{
            //    for (int y = 0; y < usched[x].Count(); y++)
            //    {
            //        for (int z = 0; z < usched[x][y].Count(); z++)
            //        {
            //            for (int w = 0; w < usched[x][y][z].Count(); w++)
            //            {
            //                ovall++;
            //                //s += usched[x][y][z][w] + " - ";
            //                if (sched.Count() > x && sched[x].Count() > y && sched[x][y].Count() > z && sched[x][y][z].Count() > w)
            //                {
            //                    //s += sched[x][y][z][w];
            //                    if (sched[x][y][z][w] == usched[x][y][z][w])
            //                        succ++;
            //                    else
            //                    {
            //                        s += Program.group.Keys.ToArray()[x] + "\t" + sched[x][y][z][w];
            //                        s += "\n";
            //                    }
            //                }
            //            }
            //        }
            //    }
            //}
            //segrid.Children.Clear();
            //var scr = new ScrollViewer();
            //var tb = new TextBlock();
            //s += Math.Round((succ * 100) / ovall, 3).ToString() + "%";
            //tb.FontSize = 40;
            //tb.Text = s;
            //scr.Content = tb;
            //segrid.Children.Add(scr);
        }

            for (int x = 0; x < usched.Count(); x++)
            {
                for (int y = 0; y < usched[x].Count(); y++)
                {
                    for (int z = usched[x][y].Count()-1; z >= 0; z--)
                    {
                        if(usched[x][y][z].Count != 0)
                        {
                            usched[x][y].RemoveRange(z+1, usched[x][y].Count - z - 1);
                            break;
                        }
                    }
                }
            }
            //MessageBox.Show((sched == usched).ToString());
            ;
            Program.schedule.schedule = usched;
            Program.schedule.Save();
            MessageBox.Show("Розклад успішно збережено.");
        }
    }
}
