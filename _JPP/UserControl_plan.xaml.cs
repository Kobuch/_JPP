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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace _JPP
{
    /// <summary>
    /// Logika interakcji dla klasy UserControl_plan.xaml
    /// </summary>
    public partial class UserControl_plan : Window
    {

        Tabelka tabelka;
        List<tabelkapokaz> tabelkapokazs;
        Tabelka_plan _tabelka_Plan { get; set; }

        private List<string> lista1 { get; set; }
        public List<string> Lista1 
        { get { return  lista1; }
            set { value = lista1; } 
 
        }

        public Tabelka_plan Tabelka_Plan
            { get { return _tabelka_Plan; }
          set { _tabelka_Plan = value; }
        }



        List<Tabelka_plan> _tabelka_Plans { get; set; }
        List<Tabelka_plan> _tabelka_Plans_tmp { get; set; }


        List<Tabelka_plan> Tabelka_Plans
        {
            get { return _tabelka_Plans; }
            set { _tabelka_Plans = value; }
        }
        List<Tabelka_plan> Tabelka_Plans_tmp
        {
            get { return _tabelka_Plans_tmp; }
            set { _tabelka_Plans_tmp = value; }
        }


        public UserControl_plan(List<tabelkapokaz> _tabelkapokazs,Tabelka _tabelka)
        {

            InitializeComponent();
            Tabelka_Plans = new List<Tabelka_plan>();
            tabelka = _tabelka;
            tabelkapokazs = _tabelkapokazs;
            Tabelka_Plan = new Tabelka_plan();

            Lista1 = new List<string>() { "text1", "text2" };
        }


        private void wklej_Click(object sender, RoutedEventArgs e)
        {
            List<string[]> wiersze = new List<string[]>();
            //Wyczysć istniejace dane  
            Tabelka_Plans.Clear();

            //sprawdz czy schowek posiada w sobie text 
            if (!Clipboard.ContainsText()) return;
            // przypisz text ze schowka do zmiennej
            string text_glowny = Clipboard.GetText();
            // podziel linie po znaku końcowym linii dodawanych z excela do danych
            string[] lines = text_glowny.Split(new[] { "\r\n", Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

            //podzielone linie podziel jeszcze raz na tablice po zmienej \t
            // i dodaj do listy ktora wyswietlasz w grid  
            foreach (string znak in lines)
            {
                string[] lines2 = znak.Split(new[] { "\t", Environment.NewLine }, StringSplitOptions.None);

                Tabelka_Plans.Add(new Tabelka_plan(lines2));
            }

            foreach (Tabelka_plan row in Tabelka_Plans)
            {
                row.Frequenz = row.Frequenz.Replace(" GHz","");
                row.Diameter = row.Diameter.Replace(",", "0,");


            }






            Grid_plan.ItemsSource = Tabelka_Plans;
            Grid_plan.Items.Refresh();




        }

        private void scal_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < Tabelka_Plans.Count; i++)
            {
                string szukana = Tabelka_Plans[i].Lfd_Nr+"a";
                if (Tabelka_Plans.Count(x => x.Lfd_Nr == szukana) > 0)
                    {
                    Tabelka_plan tmp = Tabelka_Plans.First(x => x.Lfd_Nr == szukana);
                    Tabelka_Plans[i].USER_LINK_ID += "/" + tmp.USER_LINK_ID.Substring(tmp.USER_LINK_ID.Length-4,4);
                    Tabelka_Plans[i].Ile_odu = "2";
                    Tabelka_Plans.Remove(tmp);
                    }
            }
            Grid_plan.ItemsSource = Tabelka_Plans;
            Grid_plan.Items.Refresh();




        }

        private void dopasuj_Click(object sender, RoutedEventArgs e)
        {
           // this.Stak_grid_plan.Visibility = Visibility.Collapsed;
            this.Stak_grid_dopasuj.Visibility = Visibility.Visible;
            Stak_grid_dopasuj.DataContext = Tabelka_Plans;

            this.grid_tabelka29.ItemsSource = tabelkapokazs;
            //  Grid_dopasuj.Items.Refresh();

            Tabelka_Plans_tmp = new List<Tabelka_plan>();


            foreach (tabelkapokaz row in tabelkapokazs  )
            {
               Double azym_1 =Convert.ToDouble(row.RICHTUNG.Replace(",", "."));

                int czy_jest = Tabelka_Plans.Count(x => Convert.ToDouble(x.Azimuth.Replace(",", ".")) == azym_1);

                if (czy_jest>0)
                {
                    Tabelka_plan row2 = Tabelka_Plans.First(x => Convert.ToDouble(x.Azimuth.Replace(",", ".")) == azym_1);
                    row2.Dopasuj = row.RIFU_NR;
                    Tabelka_Plans_tmp.Add(row2);
                    Tabelka_Plans.Remove(row2);
                }
                else
                {
                    Tabelka_Plans_tmp.Add(new Tabelka_plan());
                   
                }

                

            }
            
                  Tabelka_Plans_tmp.AddRange(Tabelka_Plans);

                Stak_grid_dopasuj.DataContext = Tabelka_Plans_tmp;
                

        }

        private void MyGrid_DragOver(object sender, DragEventArgs e)
        {
            DragOver_DragEnter(sender, e);

        }

        private void MyGrid_DragEnter(object sender, DragEventArgs e)
        {
            DragOver_DragEnter(sender, e);
        }
        private void DragOver_DragEnter(object sender, DragEventArgs e)
        {
            // code here to decide whether drag target is ok

            e.Effects = DragDropEffects.None;
            e.Effects = DragDropEffects.Move;
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
            return;
        }
        private void Grid_dopasuj_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {

            var dragSource = sender as DataGrid;

                Tabelka_plan data = ((DataGrid)sender).SelectedItem as Tabelka_plan;
                                 
                DragDrop.DoDragDrop(dragSource, data, DragDropEffects.Move);
            }
        }

        private void BT_akceptuj_Click(object sender, RoutedEventArgs e)
        {
            //tworzenie tabeli uzgledniającej dane z obu tabel

            for (int i = 0; i < tabelkapokazs.Count; i++)
            {
                Tabelka_plan row2 = Tabelka_Plans.First(x => x.Dopasuj== tabelkapokazs[i].RIFU_NR);
                tabelkapokazs[i]=sprawdzaj_wartosci(tabelkapokazs[i], row2);
            }
        }

        private tabelkapokaz sprawdzaj_wartosci(tabelkapokaz elem_1, Tabelka_plan elem_2)
        {
            //srednica
          
            if ( Convert.ToInt32(elem_1.RIFU.Replace(",", ".")) != Convert.ToInt32 (elem_2.Diameter.Replace(",", ".") ))  
            { elem_1.RIFU = Convert.ToInt32(elem_2.Diameter.Replace(",", ".")).ToString(); }







            return elem_1;
        }


        private void BT_zapisz_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
