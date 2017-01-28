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

        private void button_Click(object sender, RoutedEventArgs e)
        {
           Clipboard.SetText(WebControl.ExecuteJavascriptWithResult("getElementsByTagName('html')[0].innerHTML"));
        }
    }
}
