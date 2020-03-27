using Microsoft.Win32;
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

namespace XlsToCsv
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            //Konvertiere Datei
            string file = tbxPfad.Text;
            ExcelClass excel = new ExcelClass(file);
            string csvString = excel.ExcelToString();
            if (excel.StringToCsv(file, csvString))
            {
                MessageBox.Show("Konvertierung erfolgreich");
            }
            else
            {
                MessageBox.Show("Es ist ein Fehler aufgetreten\nBitte überprüfen Sie ob sich ein '.' in der Ordnerstruktur befindet", "Fehler", MessageBoxButton.OK,MessageBoxImage.Error);
            }

        }

        private void BtnAbbrechen_Click(object sender, RoutedEventArgs e)
        {
            //entferne Datei
            tbxPfad.Text = "";
        }

        private void BntSucheFile_Click(object sender, RoutedEventArgs e)
        {
            //Öffne Datei
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "New Excel Files (*.xlxs)|*.xlsx| Old Excel Files (*.xls)|*.xls";
            if (ofd.ShowDialog()==true)
            {
                tbxPfad.Text = ofd.FileName;
            }
        }

        private void TbxPfad_TextChanged(object sender, TextChangedEventArgs e)
        {
            //Aktiviere Button nur wenn eine Datei ausgewählt wurde
            btnConvert.IsEnabled = (!tbxPfad.Text.Equals(""));
        }

    }
}
