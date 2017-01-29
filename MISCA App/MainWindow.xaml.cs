using System.Windows;

namespace MISCA_App
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            get_curency();
        }

        private void again(object sender, RoutedEventArgs e)
        {
            Get_images();
        }
    }
}
