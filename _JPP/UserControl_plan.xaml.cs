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
        List<tabelkapokaz> tabelkapokazs2;
        Tabelka_plan _tabelka_Plan { get; set; }

        private List<string> lista1 { get; set; }
        public List<string> Lista1
        { get { return lista1; }
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


        public UserControl_plan(List<tabelkapokaz> _tabelkapokazs, Tabelka _tabelka)
        {

            InitializeComponent();
            Tabelka_Plans = new List<Tabelka_plan>();
            tabelka = _tabelka;
            tabelkapokazs = _tabelkapokazs;
            Tabelka_Plan = new Tabelka_plan();

            Lista1 = new List<string>() { "text1", "text2" };
        }

/// <summary>
/// WKLEJ z clipboard
///
/// </summary>
/// <param name="sender"></param>
/// <param name="e"></param>
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
                if (!string.IsNullOrEmpty(row.Frequenz)) row.Frequenz = row.Frequenz.Replace(" GHz", "");
                if (!string.IsNullOrEmpty(row.Frequenz)) row.Frequenz = row.Frequenz.Replace(" GHZ", "");

                if (!string.IsNullOrEmpty(row.Diameter)) row.Diameter = row.Diameter.Replace(",00", "0");
              

                if (!string.IsNullOrEmpty(row.Azimuth)) row.Azimuth = row.Azimuth.Replace(",", ".");

            }






            Grid_plan.ItemsSource = Tabelka_Plans;
            Grid_plan.Items.Refresh();




        }
        /// <summary>
        /// Scal - scala takie same radiolinie po numerze z literką a jako drugą
        ///  i aktualizuje liste    " Tabelka_Plans "
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void scal_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < Tabelka_Plans.Count; i++)
            {
                string szukana = Tabelka_Plans[i].Lfd_Nr + "a";
                if (Tabelka_Plans.Count(x => x.Lfd_Nr == szukana) > 0)
                {
                    Tabelka_plan tmp = Tabelka_Plans.First(x => x.Lfd_Nr == szukana);
                    Tabelka_Plans[i].USER_LINK_ID += "/" + tmp.USER_LINK_ID.Substring(tmp.USER_LINK_ID.Length - 4, 4);
                    Tabelka_Plans[i].Ile_odu = "2";
                    Tabelka_Plans.Remove(tmp);
                }
            }
            Grid_plan.ItemsSource = Tabelka_Plans;
            Grid_plan.Items.Refresh();




        }

        /// <summary>
        /// Dopasowujetabele planowane do tabeli z istniejacej.
        /// jeżeli jest taki sam to ustawia obok siebie, jeżlei nie to pod lista.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>


        private void dopasuj_Click(object sender, RoutedEventArgs e)
        {
            // this.Stak_grid_plan.Visibility = Visibility.Collapsed;
            this.Stak_grid_dopasuj.Visibility = Visibility.Visible;
            Stak_grid_dopasuj.DataContext = Tabelka_Plans;

            this.grid_tabelka29.ItemsSource = tabelkapokazs;
            //  Grid_dopasuj.Items.Refresh();


            //pusta do wypełniania
            Tabelka_Plans_tmp = new List<Tabelka_plan>();
            //pełna do uzuwania juz wykozystanych
            List<Tabelka_plan> Tabelka_Plans_tmp2=new List<Tabelka_plan>();
            Tabelka_Plans_tmp2.AddRange( Tabelka_Plans);

            int czy_jest;
            foreach (tabelkapokaz row in tabelkapokazs)
            {
                czy_jest = 0;

                if (!string.IsNullOrEmpty(row.RICHTUNG))
                {

                    Double azym_1 = Convert.ToDouble(row.RICHTUNG.Replace(",", "."));
                    czy_jest = Tabelka_Plans_tmp2.Count(x => Convert.ToDouble(x.Azimuth) == azym_1);



                    if (czy_jest > 0)
                    {
                        Tabelka_plan row2 = Tabelka_Plans_tmp2.First(x => Convert.ToDouble(x.Azimuth) == azym_1);
                        row2.Dopasuj = row.RIFU_NR;
                        Tabelka_Plans_tmp.Add(row2);
                        Tabelka_Plans_tmp2.Remove(row2);
                    }
                    else
                    {
                        Tabelka_Plans_tmp.Add(new Tabelka_plan());

                    }


                }
                else 
                { Tabelka_Plans_tmp.Add(new Tabelka_plan()); }



            }

            Tabelka_Plans_tmp.AddRange(Tabelka_Plans_tmp2);

            Stak_grid_dopasuj.DataContext = Tabelka_Plans_tmp;


        }

     


        /// <summary>
        /// Akceptuje dopasowanie i łaczy obie tabele w jedną wynikową
        /// zapisuje dane do listy "tabelkapokazs2 "
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void BT_akceptuj_Click(object sender, RoutedEventArgs e)
        {

            tabelkapokazs2 = new List<tabelkapokaz>();

            //tworzenie tabeli uzgledniającej dane z obu tabel

            List<Tabelka_plan> Tabelka_Plans_tmp3 = new List<Tabelka_plan>();
                      
            Tabelka_Plans_tmp3.AddRange(Tabelka_Plans);
            
            for (int i = 0; i < tabelkapokazs.Count; i++)
            {
               

                Tabelka_plan row2 = Tabelka_Plans.Find(x => x.Dopasuj == tabelkapokazs[i].RIFU_NR);

                if (row2 != null)
                {
                    tabelkapokazs2.Add(sprawdzaj_wartosci(tabelkapokazs[i], row2));
                    Tabelka_Plans_tmp3.Remove(row2);
                }
                else 
                {
                    tabelkapokazs2.Add(tabelkapokazs[i]);
                }
                
            }

            //dodanie tych co pozostały a ich nie było w CAD
            foreach (Tabelka_plan row2 in Tabelka_Plans_tmp3)
            {
                tabelkapokazs2.Add(sprawdzaj_wartosci_i_dodaj(row2));
            }

            this.grid_tabelka_do_zapisu.ItemsSource = tabelkapokazs2;

        }


        /// <summary>
        /// zapisuje zaktualizowanatabele do properties autocada;
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BT_zapisz_Click(object sender, RoutedEventArgs e)
        {
            if (tabelkapokazs2!=null && tabelkapokazs2.Count>0)
            {
                //zapisz 

                Obsluga_prop_cad obsluga_Prop_Cad = new Obsluga_prop_cad();
                //zapisz tabeli z properties cadowego
                obsluga_Prop_Cad.czysc_properties();
                obsluga_Prop_Cad.zapisz_tabela_aktualizujaca_planowane_do_prop_acad(tabelkapokazs2);
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Zapisano w properties \n Kolumn:  29 \n Wierszy: " + tabelkapokazs2.Count.ToString());
            }




        }


        private tabelkapokaz sprawdzaj_wartosci(tabelkapokaz elem, Tabelka_plan elem_2)
        {
            tabelkapokaz elem_1 = elem;

            //srednica

            int e1 = -1;
            int e2 = -1;
            string tmp1 = "";
         

               
             if (!string.IsNullOrEmpty(elem_1.RIFU)) tmp1 = elem_1.RIFU.Replace(",", ".");
                {
                    try
                    {

                        e1 = Convert.ToInt32(tmp1);
                        e2 = Convert.ToInt32(elem_2.Diameter);
                        if (e1 != e2) elem_1.RIFU = elem_2.Diameter;
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                }

            
             
            
    
            //czestotliwosc
            if (elem_1.FREQUENZ != elem_2.Frequenz )
            { elem_1.FREQUENZ = elem_2.Frequenz; }

            //azymut
            if (!string.IsNullOrEmpty(elem_2.Azimuth))
            {
                if (!string.IsNullOrEmpty(elem_1.RIFU) && (Convert.ToDouble(elem_1.RICHTUNG.Replace(",", ".")) != Convert.ToDouble(elem_2.Azimuth.Replace(",", "."))))
                { elem_1.RICHTUNG = Convert.ToInt32(elem_2.Azimuth.Replace(",", ".")).ToString(); }
            }

                //Gegenstelle
             if (elem_1.GEGENSTELLE != elem_2.NE_B)
                { elem_1.GEGENSTELLE = elem_2.NE_B; }

                //link
                if (elem_1.Linknummer != elem_2.USER_LINK_ID)
                { elem_1.Linknummer = elem_2.USER_LINK_ID; }

                //odu
                if (elem_1.ODU_ANZAHL != elem_2.Ile_odu || elem_2.Ile_odu=="2")
                { elem_1.ODU_ANZAHL = elem_2.Ile_odu; }

           
            return elem_1;
        }

        private tabelkapokaz sprawdzaj_wartosci_i_dodaj(Tabelka_plan elem_2)
        {
            tabelkapokaz elem_1 = new tabelkapokaz();

            elem_1.RIFU_NR = "Rifu XXX";

            elem_1.RIFU = elem_2.Diameter;

            elem_1.RICHTUNG = elem_2.Azimuth;
            elem_1.FREQUENZ = elem_2.Frequenz;
            elem_1.GEGENSTELLE = elem_2.NE_B;
            elem_1.Linknummer = elem_2.USER_LINK_ID;
            elem_1.ODU_ANZAHL = elem_2.Ile_odu;


            return elem_1;

        }




        private void Grid_dopasuj_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

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
            //if (e.LeftButton == MouseButtonState.Pressed)
            //{

            //    var dragSource = sender as DataGrid;

            //    Tabelka_plan data = ((DataGrid)sender).SelectedItem as Tabelka_plan;

            //    DragDrop.DoDragDrop(dragSource, data, DragDropEffects.Move);
            //}
        }





    }
}
