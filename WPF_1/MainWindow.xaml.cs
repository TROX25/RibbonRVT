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
using Microsoft.Win32;
using OfficeOpenXml;
using System.IO;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;

namespace WPF_1
{
    
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = ".xlsx";
            saveFileDialog.Filter = "Pliki Excel (*.xlsx)|*.xlsx|Wszystkie pliki (*.*)|*.*";

            bool? result = saveFileDialog.ShowDialog();

            if (result == true)
            {
                string filePath = saveFileDialog.FileName;

                FileInfo newFile = new FileInfo(filePath);
                if (newFile.Exists)
                {
                    newFile.Delete();  
                    newFile = new FileInfo(filePath);
                }
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    // Utwórz nowy arkusz
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Taka sama nazwa jak zestawienie w revit");
                    // Przykładowa zawartość
                    //123
                    

                    // Zapisz zmiany
                    package.Save();
                    this.Close();
                }
                MessageBox.Show("Plik Excel zapisany: " + filePath);
            }
        }

        private void AddSchedule_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("AddSchedule_Click");
        }

        private void RemoveSchedule_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("RemoveSchedule_Click");
        }
    }
}
