using System.Windows;

namespace Group4335
{
    public partial class MainWindow : Window
    {
        public MainWindow()
            => InitializeComponent();

        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            Info infow = new Info();
            infow.Show();
        }
    }
}