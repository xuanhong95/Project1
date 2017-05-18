using System.Windows;
using System.Windows.Controls;


namespace DataProcessing
{
    /// <summary>
    /// Interaction logic for Welcome.xaml
    /// </summary>
    public partial class Welcome : Page
    {
        public Welcome()
        {
            InitializeComponent();
        }

        private void gotothietlapHeso(object sender, RoutedEventArgs e)
        {
            thietlapHeSo tlhs = new thietlapHeSo();
            this.NavigationService.Navigate(tlhs);
        }

    }
}
