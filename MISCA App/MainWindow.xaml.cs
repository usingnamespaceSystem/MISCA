using System.Windows;
using System.Windows.Controls;

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

            foreach (Microsoft.Office.Interop.Excel.Worksheet sh in wbook.Worksheets)
            {
                if (!sh.Name.Contains("nul"))
                    category.Items.Add(sh.Name);
            }
        }

        
    }
}
