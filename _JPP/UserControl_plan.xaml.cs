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
    

        Tabelka_plan tabelka_Plan = new Tabelka_plan();

        List<Tabelka_plan> _tabelka_Plans { get; set; }


        List<Tabelka_plan> Tabelka_Plans
        {
            get { return _tabelka_Plans; }
            set { _tabelka_Plans = value; }


        }

        public UserControl_plan(Tabelka tabelka)
        {
            InitializeComponent();
            Tabelka_Plans = new List<Tabelka_plan>();

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
                    Tabelka_Plans.Remove(tmp);
                    }
            }
            Grid_plan.ItemsSource = Tabelka_Plans;
            Grid_plan.Items.Refresh();




        }

        private void dopasuj_Click(object sender, RoutedEventArgs e)
        {

        }








    }
}
