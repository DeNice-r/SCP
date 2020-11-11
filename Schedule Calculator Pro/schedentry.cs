using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Schedule_Calculator_Pro
{
    internal class Schedentry : Border
    {
        public static List<List<int>> colwidth = new List<List<int>>();
        public int[] idx = { -1, -1, -1 };
        public string[] info = { "", "", "" };
        private static BitmapImage cutimg = new BitmapImage(new Uri("Images\\cut.png", UriKind.Relative));
        private static BitmapImage insimg = new BitmapImage(new Uri("Images\\insert.png", UriKind.Relative));
        public DockPanel dock = new DockPanel();
        public StackPanel subjects = new StackPanel();
        public StackPanel dons = new StackPanel();
        public StackPanel auds = new StackPanel();
        public StackPanel ctrls = new StackPanel();
        public Label[] subj = new Label[] { new Label(), new Label() };
        public Label[] don = new Label[] { new Label(), new Label() };
        public Label[] aud = new Label[] { new Label(), new Label() };
        public Border donbord = new Border();
        public static Schedentry cut = null;
        public static int buttonsizes = 25;

        public Schedentry()
        {
            construct(info, idx);
        }

        public Schedentry(string[] infos)
        {
            construct(infos, idx);
        }

        public Schedentry(string[] infos, int[] idxs)
        {
            construct(infos, idxs);
        }

        private void construct(string[] infos, int[] idxs)
        {
            idx = idxs;
            info = infos;
            donbord.BorderBrush = Brushes.Black; donbord.Child = dons;

            if (infos.Length >= 3 || infos.Length == 0)
            {
                subj[0].Width = colwidth[idxs[0]][0]; subj[1].Width = colwidth[idxs[0]][0];
                don[0].Width = colwidth[idxs[0]][1]; don[1].Width = colwidth[idxs[0]][1];
                aud[0].Width = colwidth[idxs[0]][2]; aud[1].Width = colwidth[idxs[0]][2];
            }

            switch (infos.Length)
            {
                case 1:
                    {
                        subj[0].Content = infos[0]; subj[0].VerticalContentAlignment = VerticalAlignment.Center;
                        subj[0].Width = colwidth[idxs[0]][0] + colwidth[idxs[0]][1] + colwidth[idxs[0]][2];

                        if (idxs.Length == 1)
                            subj[0].FontSize = 20;
                    }
                    break;

                case 3:
                    {
                        subj[0].Content = infos[0]; subj[0].Height = 50; subj[0].VerticalContentAlignment = VerticalAlignment.Center;
                        don[0].Content = infos[1]; don[0].Height = 50; don[0].VerticalContentAlignment = VerticalAlignment.Center;
                        aud[0].Content = infos[2]; aud[0].Height = 50; aud[0].VerticalContentAlignment = VerticalAlignment.Center;
                    }
                    break;

                case 5:
                    {
                        subj[0].Content = infos[0]; subj[0].Height = 50; subj[0].VerticalContentAlignment = VerticalAlignment.Center;
                        don[0].Content = infos[1]; don[1].Content = infos[2];
                        aud[0].Content = infos[3]; aud[1].Content = infos[4];
                        don[0].Height = 25; aud[0].Height = 25;
                    }
                    break;

                case 6:
                    {
                        subj[0].Content = infos[0]; subj[1].Content = infos[1];
                        don[0].Content = infos[2]; don[1].Content = infos[3];
                        aud[0].Content = infos[4]; aud[1].Content = infos[5];
                        subj[0].Height = 25; don[0].Height = 25; aud[0].Height = 25;
                    }
                    break;

                default:
                    break;
            }
            subjects.Children.Add(subj[0]); subjects.Children.Add(subj[1]);
            dons.Children.Add(don[0]); dons.Children.Add(don[1]);
            auds.Children.Add(aud[0]); auds.Children.Add(aud[1]);
            if (idxs.Length == 3)
            {
                donbord.BorderThickness = new Thickness(1, 0, 1, 0);
                Image cutimage = new Image(); cutimage.Source = cutimg;
                Image insimage = new Image(); insimage.Source = insimg;
                var ctrlpanel = new StackPanel(); var ctrlbord = new Border(); ctrlbord.BorderThickness = new Thickness(1, 0, 0, 0); ctrlbord.BorderBrush = Brushes.Black;
                ctrlbord.Child = ctrlpanel;
                var cutbutt = new Buttonplus(); cutbutt.Click += SubjCut_Click; cutbutt.idx = idxs; ctrlpanel.Children.Add(cutbutt);
                cutbutt.Content = cutimage; cutbutt.Height = cutbutt.Width = buttonsizes;
                var insbutt = new Buttonplus(); insbutt.Click += SubjIns_Click; insbutt.idx = idxs; ctrlpanel.Children.Add(insbutt);
                insbutt.Content = insimage; insbutt.Height = insbutt.Width = buttonsizes;
                ctrls.Children.Add(ctrlbord);
            }
            this.BorderBrush = Brushes.Black;
            this.BorderThickness = new Thickness(1);
            this.MouseLeftButtonDown += Cpl_DragEnter;
            dock.Children.Add(subjects); dock.Children.Add(donbord); dock.Children.Add(auds); dock.Children.Add(ctrls);
            dock.Height = 50;
            this.Child = dock;
        }

        public void resize(int s = -1, int d = -1, int a = -1)
        {
            if (s != -1)
                ((Label)subjects.Children[0]).Width = ((Label)subjects.Children[1]).Width = s;
            if (d != -1)
                ((Label)dons.Children[0]).Width = ((Label)dons.Children[1]).Width = d;
            if (a != -1)
                ((Label)auds.Children[0]).Width = ((Label)auds.Children[1]).Width = a;
        }

        public void resize(int[] wths)
        {
            if (wths[0] != -1)
                ((Label)subjects.Children[0]).Width = ((Label)subjects.Children[1]).Width = wths[0];
            if (wths[1] != -1)
                ((Label)dons.Children[0]).Width = ((Label)dons.Children[1]).Width = wths[1];
            if (wths[2] != -1)
                ((Label)auds.Children[0]).Width = ((Label)auds.Children[1]).Width = wths[2];
        }

        private void SubjCut_Click(object sender, RoutedEventArgs e)
        {
            ScheduleEditor.deselect();
            var s = (Buttonplus)sender;
            cut = (Schedentry)((DockPanel)((StackPanel)((Border)((StackPanel)s.Parent).Parent).Parent).Parent).Parent;
            cut.Opacity = .35;
        }

        private void SubjIns_Click(object sender, RoutedEventArgs e)
        {
            var s = (Buttonplus)sender;
            var sentry = (Schedentry)((DockPanel)((StackPanel)((Border)((StackPanel)s.Parent).Parent).Parent).Parent).Parent;
            if (sentry == cut || cut == null)
            {
                ScheduleEditor.deselect();
                return;
            }

            var from = (StackPanel)cut.Parent;
            var to = (StackPanel)sentry.Parent;

            int[] w1 = colwidth[Program.group.Keys.ToList().IndexOf(((Schedentry)from.Children[0]).info[0])].ToArray();
            int[] w2 = colwidth[Program.group.Keys.ToList().IndexOf(((Schedentry)to.Children[0]).info[0])].ToArray();
            int[] mw;

            if (w1 != w2)
            {
                mw = new int[3] { max(w1[0], w2[0]), max(w1[1], w2[1]), max(w1[2], w2[2]) };

                foreach (Schedentry f in from.Children)
                {
                    f.resize(mw);
                }
                foreach (Schedentry t in to.Children)
                {
                    t.resize(mw);
                }
            }

            var index1 = from.Children.IndexOf(cut);
            var index2 = to.Children.IndexOf(sentry);

            var temp = sentry.idx;
            sentry.idx = cut.idx;
            cut.idx = temp;

            to.Children.Remove(sentry);
            from.Children.Remove(cut);

            if (index1 > index2)
            {
                to.Children.Insert(index2, cut);
                from.Children.Insert(index1, sentry);
            }
            else
            {
                from.Children.Insert(index1, sentry);
                to.Children.Insert(index2, cut);
            }

            ScheduleEditor.deselect();
        }

        private void Cpl_DragEnter(object sender, MouseButtonEventArgs e)
        {
        }

        private int max(int e1, int e2)
        {
            return (e1 > e2) ? e1 : e2;
        }
    }
}