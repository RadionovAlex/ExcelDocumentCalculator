using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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

namespace ExcelDocumentCalculator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string _templateFilePath = string.Empty;
        public MainWindow()
        {
            InitializeComponent();

            btn_chooseFile.IsEnabled = false;
        }

        private void ChooseFile_Click (object sender, RoutedEventArgs e)
        {
            if (!ValidateInput(out var maxRows, out var minInvoice, out var maxInvoice, out var hourRate))
            {
                ShowAlert("Please, make correct inout values first");
                return;
            }
                

            // Create an instance of OpenFileDialog
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "My love, select .xlsx File to calculate your wonderful work hours",
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*" // File type filter
            };

            // Show the dialog and check if the user selected a file
            if (openFileDialog.ShowDialog() == true)
            {
                string selectedFilePath = openFileDialog.FileName; // Get the full file path
                MessageBox.Show($"You selected: {selectedFilePath}", "File Selected");

                // Use this file path in your application
                // Example: Pass it to your processing method
                ProcessFile(selectedFilePath, maxRows, minInvoice, maxInvoice, hourRate);
            }
        }

        private void ChooseTemplate_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";  // You can change the filter as per the file types you're working with

            if (openFileDialog.ShowDialog() == true)
            {
                _templateFilePath = openFileDialog.FileName;
                btn_chooseFile.IsEnabled = true;
                ShowMessage("Continue");
            }
        }

        private bool ValidateInput(out int maxRows, out int minInvoiceVolume, out int maxInvoiceVolume, out float hourRate)
        {
            maxRows = 0;
            minInvoiceVolume = 0;
            maxInvoiceVolume = 0;
            hourRate = 0;

            if (!Int32.TryParse(txtMaxRowsInInvoice.Text, out var rowsValue) || rowsValue > 10)
            {
                ShowAlert("Cannot parse max Rows In Invoice to number or it`s value more than 10");
                return false;
            }
            maxRows = rowsValue;

            if(!Int32.TryParse(txtMinInvoiceVolume.Text, out var minInvoice))
            {
                ShowAlert("Cannot parse min invoice volume to number");
                return false;
            }
            minInvoiceVolume = minInvoice;

            if (!Int32.TryParse(txtMaxInvoiceVolume.Text, out var maxInvoice))
            {
                ShowAlert("Cannot parse min invoice volume to number");
                return false;
            }
            maxInvoiceVolume = maxInvoice;

            if (!float.TryParse(txtHourRate.Text, out var ratePerHour))
            {
                ShowAlert("Cannot parse min invoice volume to number");
                return false;
            }
            hourRate = ratePerHour;


            if(maxInvoiceVolume < minInvoiceVolume)
            {
                ShowAlert("Max invoice volume is less than Min invoice voulume");
                return false;
            }

            return true;
        }

        private void ProcessFile(string filePath, int rowsCount, int minInvoice, int maxInvoice, float hourRate)
        {
            var dirPath = Path.GetDirectoryName(filePath);
            new HoursCalculatorSimple().Calculate(filePath, _templateFilePath, () => OpenWindowsExplorer(dirPath),rowsCount, minInvoice, maxInvoice, hourRate);
            MessageBox.Show($"Processing file: {filePath}");
        }

        static void OpenWindowsExplorer(string path)
        {
            if (!string.IsNullOrWhiteSpace(path))
            {                
                Process.Start("explorer.exe", path);
            }
        }

        // Method to show alert messages
        private void ShowAlert(string message)
        {
            MessageBox.Show(message, "Not valid input", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void ShowMessage(string message)
        {
            MessageBox.Show(message, "Good, now let`s load working hours file", MessageBoxButton.OK);
        }
    }
}
