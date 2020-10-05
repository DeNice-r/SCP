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
    class schedentry : Border
    {
        public static List<List<uint>> colwidth = new List<List<uint>>();
        public static List<List<List<List<string>>>> sched;
        private static BitmapImage cutimg = new BitmapImage(new Uri("Images\\cut.png", UriKind.Relative));
        private static BitmapImage insimg = new BitmapImage(new Uri("Images\\insert.png", UriKind.Relative));
        DockPanel dock = new DockPanel();
        StackPanel subjects = new StackPanel();
        StackPanel dons = new StackPanel();
        StackPanel auds = new StackPanel();
        StackPanel ctrls = new StackPanel();
        int fontsize = 12;

        public schedentry(string[] info, int[] idx)
        {
            Label[] subj = new Label[] { new Label(), new Label() };
            Label[] don = new Label[] { new Label(), new Label() };
            Label[] aud = new Label[] { new Label(), new Label() };

            subj[0].Width = colwidth[idx[0]][0]; subj[1].Width = colwidth[idx[0]][0];
            don[0].Width = colwidth[idx[0]][1]; don[1].Width = colwidth[idx[0]][1];
            aud[0].Width = colwidth[idx[0]][2]; aud[1].Width = colwidth[idx[0]][2];

            switch (info.Length)
            {
                case 1:
                    {
                        subj[0].Width = colwidth[idx[0]][0];
                        don[0].Width = colwidth[idx[0]][1];
                        aud[0].Width = colwidth[idx[0]][2];
                    }
                    break;
                case 3:
                    {
                        subj[0].Content = info[0]; subj[0].Height = 50; subj[0].VerticalContentAlignment = VerticalAlignment.Center;
                        don[0].Content = info[1]; don[0].Height = 50; don[0].VerticalContentAlignment = VerticalAlignment.Center;
                        aud[0].Content = info[2]; aud[0].Height = 50; aud[0].VerticalContentAlignment = VerticalAlignment.Center;
                    }
                    break;
                case 5:
                    {
                        subj[0].Content = info[0]; subj[0].Height = 50; subj[0].VerticalContentAlignment = VerticalAlignment.Center;
                        don[0].Content = info[1]; don[1].Content = info[2];
                        aud[0].Content = info[3]; aud[1].Content = info[4];
                        don[0].Height = 25; aud[0].Height = 25;
                    }
                    break;
                case 6:
                    {
                        subj[0].Content = info[0]; subj[1].Content = info[1];
                        don[0].Content = info[2]; don[1].Content = info[3];
                        aud[0].Content = info[4]; aud[1].Content = info[5];
                        subj[0].Height = 25; don[0].Height = 25; aud[0].Height = 25;
                    }
                    break;
                default:
                    break;
            }
            subjects.Children.Add(subj[0]); subjects.Children.Add(subj[1]);
            dons.Children.Add(don[0]); dons.Children.Add(don[1]);
            auds.Children.Add(aud[0]); auds.Children.Add(aud[1]);
            Image cutimage = new Image(); cutimage.Source = cutimg;
            Image insimage = new Image(); insimage.Source = insimg;
            var ctrlpanel = new StackPanel(); var ctrlbord = new Border(); ctrlbord.BorderThickness = new Thickness(1, 0, 0, 0); ctrlbord.BorderBrush = Brushes.Black;
            ctrlbord.Child = ctrlpanel;
            var cutbutt = new Buttonplus(); cutbutt.Click += SubjCut_Click; cutbutt.idx = idx[0] + "." + idx[1]+ "." + idx[2]; ctrlpanel.Children.Add(cutbutt);
            cutbutt.Content = cutimage; cutbutt.Height = cutbutt.Width = 25;
            var insbutt = new Buttonplus(); insbutt.Click += SubjIns_Click; insbutt.idx = idx[0] + "." + idx[1]+ "." + idx[2]; ctrlpanel.Children.Add(insbutt);
            insbutt.Content = insimage; insbutt.Height = insbutt.Width = 25;
            ctrls.Children.Add(ctrlbord);
            this.BorderBrush = Brushes.Black;
            this.BorderThickness = new Thickness(1);
            this.MouseLeftButtonDown += Cpl_DragEnter;
            dock.Children.Add(subjects); dock.Children.Add(dons); dock.Children.Add(auds); dock.Children.Add(ctrls);
            dock.Height = 50;
            this.Child = dock;
        }

        private void SubjCut_Click(object sender, RoutedEventArgs e)
        {
            var s = (Buttonplus)sender;
            var idx = s.idx.Split('.');
            var curcpl = sched[Convert.ToInt32(idx[0])][Convert.ToInt32(idx[1])][Convert.ToInt32(idx[2])];
            string show = "";
            foreach (var str in curcpl)
                show += str + " ";
            MessageBox.Show(show);
        }

        private void SubjIns_Click(object sender, RoutedEventArgs e)
        {
            var s = (Buttonplus)sender;
            var idx = s.idx.Split('.');
            var curcpl = sched[Convert.ToInt32(idx[0])][Convert.ToInt32(idx[1])][Convert.ToInt32(idx[2])];
            string show = "";
            foreach (var str in curcpl)
                show += str + " ";
            MessageBox.Show(show);
        }

        private void Cpl_DragEnter(object sender, MouseButtonEventArgs e)
        {
            var dp = (DockPanel)sender;
            //MessageBox.Show("");
            //if()
        }
    }
}

