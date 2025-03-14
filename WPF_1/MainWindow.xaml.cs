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
using Path = System.IO.Path;

namespace WPF_1
{

    public partial class MainWindow : Window
    {
        private UIApplication _uiapp;
        public MainWindow(UIApplication uiapp)
        {
            InitializeComponent();
            _uiapp = uiapp;
            DodanieZestawien(uiapp); // pobieram od razu liste przy włączeniu okna wpf
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = ".xlsx";
            saveFileDialog.Filter = "Pliki Excel (*.xlsx)|*.xlsx|Wszystkie pliki (*.*)|*.*";

            bool? result = saveFileDialog.ShowDialog();
            if (result == true)
            {
                string excelFilePath = saveFileDialog.FileName;
                if (!excelFilePath.EndsWith(".xlsx"))
                {
                    excelFilePath += ".xlsx";
                }

                Document doc = _uiapp.ActiveUIDocument.Document;
                string tempCsvPath = Path.Combine(Path.GetTempPath(), "temp_schedule.csv");

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage())
                {
                    foreach (string scheduleName in Zestawienia_Excel.Items)
                    {
                        ViewSchedule schedule = new FilteredElementCollector(doc)
                            .OfClass(typeof(ViewSchedule))
                            .Cast<ViewSchedule>()
                            .FirstOrDefault(s => s.Name == scheduleName);

                        if (schedule != null)
                        {
                            // Eksportujemy zestawienie do CSV
                            ViewScheduleExportOptions options = new ViewScheduleExportOptions();
                            options.FieldDelimiter = ","; 
                            

                            schedule.Export(Path.GetTempPath(), "temp_schedule.csv", options);

                            // Wczytujemy CSV do Excela
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(schedule.Name);
                            string[] csvLines = File.ReadAllLines(tempCsvPath);
                            for (int row = 0; row < csvLines.Length; row++)
                            {
                                string[] cells = csvLines[row].Split(',');
                                for (int col = 0; col < cells.Length; col++)
                                {
                                    worksheet.Cells[row + 1, col + 1].Value = cells[col].Trim('"'); // Usunięcie cudzysłowów
                                }
                            }
                        }
                    }

                    package.SaveAs(new FileInfo(excelFilePath));
                }
                this.Close();
                MessageBox.Show("Plik Excel zapisany: " + excelFilePath);
            }
        }

        private void AddSchedule_Click(object sender, RoutedEventArgs e) // Opis działania przycisku dodania zestawienia do exportu
        {
            if (Zestawienia_Revit.SelectedItem != null)
            {
                string selectedSchedule = Zestawienia_Revit.SelectedItem.ToString();
                if (!Zestawienia_Excel.Items.Contains(selectedSchedule))
                {
                    Zestawienia_Excel.Items.Add(Zestawienia_Revit.SelectedItem); // !!! KOLEJNOŚĆ MA ZNACZENIE, NIE PRZECHOWUJEMY OZNAKOWANYCH ZESTAWIEŃ W INNYM MIEJSCU
                    Zestawienia_Revit.Items.Remove(Zestawienia_Revit.SelectedItem);
                }
            }
        }
        
        private void RemoveSchedule_Click(object sender, RoutedEventArgs e) // Opis działania przycisku usunięcia zestawienia z exportu
        {
            if (Zestawienia_Excel.SelectedItem != null)
            {
                Zestawienia_Revit.Items.Add(Zestawienia_Excel.SelectedItem);   // !!! KOLEJNOŚĆ MA ZNACZENIE, NIE PRZECHOWUJEMY OZNAKOWANYCH ZESTAWIEŃ W INNYM MIEJSCU
                Zestawienia_Excel.Items.Remove(Zestawienia_Excel.SelectedItem);
            }
        }
                            // ALTERNATYWA 
        /*
        private void RemoveSchedule_Click(object sender, RoutedEventArgs e)
        {
            if (Zestawienia_Excel.SelectedItem != null)
            {
                var selectedItem = Zestawienia_Excel.SelectedItem;

                Zestawienia_Excel.Items.Remove(selectedItem);

                Zestawienia_Revit.Items.Add(selectedItem);
            }
        }
        */

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
    

