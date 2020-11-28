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
    /// Logika interakcji dla klasy UserControl2.xaml
    /// </summary>
    public partial class UserControl2 : Window
    
      {
        private List<tabelkapokaz20> tabelkapokazs20;
        private List<tabelkapokaz> tabelkapokazs;
       
        Tabelka tabelka = new Tabelka();
        Obsluga_prop_cad obsluga_Prop_Cad = new Obsluga_prop_cad();

        public UserControl2(  List<tabelkapokaz20> tabelkapokazs20, Tabelka tabelka)
        {
            InitializeComponent();
            this.tabelkapokazs20 = tabelkapokazs20;
            this.tabelka = tabelka;
           
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataGrid20.ItemsSource = tabelkapokazs20;
            TBilosckolumn.Text = tabelka.ilekolumn.ToString();
            TBiloscwierszy.Text = tabelka.ilewierszy.ToString();
            TBkierpulnocy.Text = tabelka.kierpolnocy_deg;
        



        }

        private void zmien20_na29_Click(object sender, RoutedEventArgs e)
        {
            
            tabelkapokazs = obsluga_Prop_Cad.przerobtabelepokaz20_na_29(tabelkapokazs20);

            dataGrid29.ItemsSource = tabelkapokazs;
           

        }

        private void wykonaj20_na29_Click(object sender, RoutedEventArgs e)
        {
           
            tabelkapokazs = obsluga_Prop_Cad.przerobtabelepokaz20_na_29(tabelkapokazs20);
            obsluga_Prop_Cad.przerobtabelepokaz20_na_29_zapisz(tabelkapokazs);
        }
    }
}
