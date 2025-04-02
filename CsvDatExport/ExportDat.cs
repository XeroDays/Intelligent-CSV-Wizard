using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

namespace CsvDatExport
{
    public partial class ExportDat : Form
    {
        public ExportDat()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Data Files (*.dat;*.csv)|*.dat;*.csv|All Files (*.*)|*.*",
                Title = "Select a File"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtFilePath.Text = openFileDialog.FileName;
            }
        }

        private void ProcessButton_Click(object sender, EventArgs e)
        {
            string filePath = txtFilePath.Text;
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                MessageBox.Show("Please select a valid file.", "Error");
                return;
            }

            try
            {
                string[] lines = File.ReadAllLines(filePath);
                int maxColumns = 0; // Track maximum columns

                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    Title = "Save Excel File",
                    FileName = "ExportedData.xlsx"
                };

                if (saveFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                string excelFilePath = saveFileDialog.FileName;
                var delimiter = txtDelimiter.Text.Trim().FirstOrDefault();

                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Data");

                    for (int i = 0; i < lines.Length; i++)
                    {
                        string[] columns = lines[i].Split(delimiter);
                        maxColumns = Math.Max(maxColumns, columns.Length); // Update max columns

                        for (int j = 0; j < columns.Length; j++)
                        {
                            worksheet.Cells[i + 1, j + 1].Value = columns[j].Trim();
                        }
                    }

                    package.SaveAs(new FileInfo(excelFilePath));
                }

                // Show summary message
                MessageBox.Show(
                    $"Successfully exported {lines.Length} rows and {maxColumns} columns.",
                    "Export Complete",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error");
            }
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            ProcessButton_Click(sender, e);
        }
    }
}
