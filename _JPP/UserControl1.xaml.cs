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
using static _JPP.HKT_class;

namespace _JPP
{
    /// <summary>
    /// Logika interakcji dla klasy UserControl1.xaml
    /// </summary>
    public partial class UserControl1 : Window
    {
        public List<tabelkapokaz> tabelkapokazs;
        Tabelka tabelka = new Tabelka();
        public UserControl1(List<tabelkapokaz> tabelkapokazs, Tabelka tabelka)
        {
            InitializeComponent();
            this.tabelkapokazs = tabelkapokazs;
            this.tabelka = tabelka;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            dataGrid.ItemsSource = tabelkapokazs;
            TBilosckolumn.Text = tabelka.ilekolumn.ToString();
            TBiloscwierszy.Text = tabelka.ilewierszy.ToString();
            TBkierpulnocy.Text = tabelka.kierpolnocy_deg;

                
                
        }
    }
}
