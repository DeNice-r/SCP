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

            for (int _group = 0; _group < sched.Count(); _group++)
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
                //

                var border2 = new Border();
                border2.BorderBrush = Brushes.Black;
                border2.BorderThickness = new Thickness(1);
                var label1 = new Label();
                border2.Child = label1;
                label1.Content = Program.group.Keys.ToArray()[_group];
                label1.HorizontalContentAlignment = HorizontalAlignment.Center;
                label1.FontSize = 20;

                uigroups.Add(new StackPanel());
                uigroups[_group].Children.Add(border2);
                for (int _day = 0; _day < 5; _day++)
                {
                    var label = new Label(); // temp. textblock
                    label.Content = DOW[_day];
                    label.HorizontalContentAlignment = HorizontalAlignment.Center;
                    var daybord = new Border();
                    daybord.BorderBrush = Brushes.Black;
                    daybord.BorderThickness = new Thickness(1);
                    daybord.Child = label;

                    uigroups[_group].Children.Add(daybord);
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
                                var subjpanel = new StackPanel();
                                var donpanel = new StackPanel(); var donbord = new Border(); donbord.BorderBrush = Brushes.Black; donbord.BorderThickness = new Thickness(1, 0, 1, 0);
                                donbord.Child = donpanel;
                                var audpanel = new StackPanel();
                                var ctrlpanel = new StackPanel();

                                var tbsubj1 = new Label(); subjpanel.Children.Add(tbsubj1);
                                var tbsubj2 = new Label(); subjpanel.Children.Add(tbsubj2);
                                var tbdon1 = new Label(); donpanel.Children.Add(tbdon1);
                                var tbdon2 = new Label(); donpanel.Children.Add(tbdon2);
                                var tbaud1 = new Label(); audpanel.Children.Add(tbaud1);
                                var tbaud2 = new Label(); audpanel.Children.Add(tbaud2);
                                tbsubj1.Width = tbsubj2.Width = lensubj; tbdon1.Width = tbdon2.Width = lendon; tbaud1.Width = tbaud2.Width = lenaud;

                                if (curlen == 5)
                                {
                                    tbsubj1.Content = curcpl[0]; tbsubj1.Height = 50; tbsubj1.VerticalContentAlignment = VerticalAlignment.Center;
                                    tbdon1.Content = curcpl[1]; tbdon2.Content = curcpl[2];
                                    tbaud1.Content = curcpl[3]; tbaud2.Content = curcpl[4];
                                    tbdon1.Height = 25; tbaud1.Height = 25;
                                }
                                else if (curlen == 6)
                                {
                                    tbsubj1.Content = curcpl[0]; tbsubj2.Content = curcpl[1];
                                    tbdon1.Content = curcpl[2]; tbdon2.Content = curcpl[3];
                                    tbaud1.Content = curcpl[4]; tbaud2.Content = curcpl[5];
                                    tbsubj1.Height = 25; tbdon1.Height = 25; tbaud1.Height = 25;
                                }
                                else
                                {
                                    tbsubj1.Content = curcpl[0]; tbsubj1.Width = lensubj; tbsubj1.Height = 50; tbsubj1.VerticalContentAlignment = VerticalAlignment.Center;
                                    tbdon1.Content = curcpl[1]; tbdon1.Height = 50; tbdon1.VerticalContentAlignment = VerticalAlignment.Center;
                                    tbaud1.Content = curcpl[2]; tbaud1.Height = 50; tbaud1.VerticalContentAlignment = VerticalAlignment.Center;
                                }

                                dp.Children.Add(subjpanel); dp.Children.Add(donbord); dp.Children.Add(audpanel); dp.Children.Add(ctrlpanel);
                            }
                        }
                        dp.MouseLeftButtonDown += Cpl_DragEnter;
                        dp.Height = 50;
                        uigroups[_group].Children.Add(border);
                    }
                }
            }
            foreach (var group in uigroups)
                docks.Children.Add(group);
        } // Перетягивание на коллбеке дрега с переводом сендера в текстблок.

        private void Window_Closed(object sender, EventArgs e)
        {
            Program.scheditwin = null;
            if (Program.workwithschedit)
            {
                if (parent != null)
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
        } }
}
