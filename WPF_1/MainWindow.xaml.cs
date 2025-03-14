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
        public MainWindow(UIApplication uiapp)
        {
            InitializeComponent();
            DodanieZestawien(uiapp);
        }

        private void Button_Click(object sender, RoutedEventArgs e) // Opis działania Głównego przycisku eksportu
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();   // Okno dialogowe zapisu pliku, czyli eksplorator plików
            saveFileDialog.DefaultExt = ".xlsx";    // Defaultowy format pliku do zapisu to .xlsx
            saveFileDialog.Filter = "Pliki Excel (*.xlsx)|*.xlsx|Wszystkie pliki (*.*)|*.*"; // Wyśwetla tylko pliki Excela

            bool? result = saveFileDialog.ShowDialog(); // W

            if (result == true)
            {
                string filePath = saveFileDialog.FileName;

                FileInfo newFile = new FileInfo(filePath);
                if (newFile.Exists)
                {
                    newFile.Delete();
                    newFile = new FileInfo(filePath);
                }
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial; // Wykorzystanie darmowej wersji biblioteki do zapisu pliku Excel
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
                MessageBox.Show("Plik Excel zapisany: " + filePath);   // Komunikat o zapisaniu pliku
            }
        }

        private void AddSchedule_Click(object sender, RoutedEventArgs e) // Opis działania przycisku dodania zestawienia do exportu
        {
            if (Zestawienia_Revit.SelectedItem != null)
            {
                string selectedSchedule = Zestawienia_Revit.SelectedItem.ToString();
                if (!Zestawienia_Excel.Items.Contains(selectedSchedule))
                {
                    Zestawienia_Excel.Items.Add(selectedSchedule);
                }
            }
        }

        private void RemoveSchedule_Click(object sender, RoutedEventArgs e) // Opis działania przycisku usunięcia zestawienia z exportu
        {
            if (Zestawienia_Excel.SelectedItem != null)
            {
                Zestawienia_Excel.Items.Remove(Zestawienia_Excel.SelectedItem);
            }
        }

        private void DodanieZestawien(UIApplication uiapp) // Funkcja pobierająca zestawienia z Revita
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Pobranie zestawień z Revita
            var schedules = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .ToList();

            // Przypisanie zestawień do ListBoxa, pomijając "Revision Schedule"
            foreach (var schedule in schedules)
            {
                // Sprawdzamy, czy nazwa zawiera "Revision Schedule", jeśli tak to nie jest wyświetlana w oknie dostępych zestawień
                if (!schedule.Name.Contains("Revision Schedule"))
                {
                    Zestawienia_Revit.Items.Add(schedule.Name);
                }
            }

            // Ustawienie trybu wyboru tylko jednego elementu
            Zestawienia_Revit.SelectionMode = SelectionMode.Single;

            // Obsługa zmiany zaznaczenia
            Zestawienia_Revit.SelectionChanged += (sender, e) =>
            {
                if (Zestawienia_Revit.SelectedItem != null)
                {
                    // Jeśli coś zostało wybrane, odznacz inne (w tym przypadku jedno zaznaczenie jest dozwolone)
                    var selectedItem = Zestawienia_Revit.SelectedItem.ToString();
                    Console.WriteLine($"Zaznaczone zestawienie: {selectedItem}");
                 
                }
            };
        }
    }
}
    

