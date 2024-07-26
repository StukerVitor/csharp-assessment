using Parquet.Schema;
using Parquet;
using System.Data;
using OxyPlot;
using OxyPlot.Series;
using OxyPlot.WindowsForms;
using DataColumn = System.Data.DataColumn;
using OxyPlot.Axes;
using System.Text;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DataField = Parquet.Schema.DataField;

namespace ArcVera_Tech_Test
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private async void btnImportEra5_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Parquet files (*.parquet)|*.parquet|All files (*.*)|*.*";
                openFileDialog.Title = "Select a Parquet File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    DataTable dataTable = await ReadParquetFileAsync(filePath);
                    dgImportedEra5.DataSource = dataTable;
                    PlotU10DailyValues(dataTable);
                }
            }
        }

        private async Task<DataTable> ReadParquetFileAsync(string filePath)
        {
            using (Stream fileStream = File.OpenRead(filePath))
            {
                using (var parquetReader = await ParquetReader.CreateAsync(fileStream))
                {
                    DataTable dataTable = new DataTable();

                    for (int i = 0; i < parquetReader.RowGroupCount; i++)
                    {
                        using (ParquetRowGroupReader groupReader = parquetReader.OpenRowGroupReader(i))
                        {
                            // Create columns
                            foreach (DataField field in parquetReader.Schema.GetDataFields())
                            {
                                if (!dataTable.Columns.Contains(field.Name))
                                {
                                    Type columnType = field.HasNulls ? typeof(object) : field.ClrType;
                                    dataTable.Columns.Add(field.Name, columnType);
                                }

                                // Read values from Parquet column
                                DataColumn column = dataTable.Columns[field.Name];
                                Array values = (await groupReader.ReadColumnAsync(field)).Data;
                                for (int j = 0; j < values.Length; j++)
                                {
                                    if (dataTable.Rows.Count <= j)
                                    {
                                        dataTable.Rows.Add(dataTable.NewRow());
                                    }
                                    dataTable.Rows[j][field.Name] = values.GetValue(j);
                                }
                            }
                        }
                    }

                    return dataTable;
                }
            }
        }

        private void PlotU10DailyValues(DataTable dataTable)
        {
            var plotModel = new PlotModel { Title = "Daily u10 Values" };
            var lineSeries = new LineSeries { Title = "u10" };

            var groupedData = dataTable.AsEnumerable()
                .GroupBy(row => DateTime.Parse(row["date"].ToString()))
                .Select(g => new
                {
                    Date = g.Key,
                    U10Average = g.Average(row => Convert.ToDouble(row["u10"]))
                })
                .OrderBy(data => data.Date);

            foreach (var data in groupedData)
            {
                lineSeries.Points.Add(new DataPoint(DateTimeAxis.ToDouble(data.Date), data.U10Average));
            }

            plotModel.Series.Add(lineSeries);
            plotView1.Model = plotModel;
        }

        private void btnExportCsv_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                saveFileDialog.Title = "Save CSV File";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;
                    DataTable dataTable = (DataTable)dgImportedEra5.DataSource;

                    if (dataTable != null)
                    {
                        try
                        {
                            StringBuilder csvContent = new StringBuilder();

                            // Add the column headers (Columns.Cast<DataColumn>() to convert Columns(DataColumnCollection) to IEnumerable<DataColumn> for LINQ operations)
                            IEnumerable<string> columnNames = dataTable.Columns.Cast<DataColumn>()
                                                     .Select(column => column.ColumnName);
                            csvContent.AppendLine(string.Join(",", columnNames));

                            // Add the rows
                            foreach (DataRow row in dataTable.Rows)
                            {
                                // Double single quotes to avoid confusion during csv export
                                IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString().Replace("\"", "\"\""));
                                csvContent.AppendLine(string.Join(",", fields));
                            }

                            // Write to file
                            File.WriteAllText(filePath, csvContent.ToString());
                            MessageBox.Show("CSV file has been saved successfully.", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"An error occurred while saving the CSV file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No data to export.", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveFileDialog.Title = "Save Excel File";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;
                    DataTable dataTable = (DataTable)dgImportedEra5.DataSource;

                    if (dataTable != null)
                    {
                        try
                        {
                            using (var workbook = new XLWorkbook())
                            {
                                const int maxRowsPerSheet = 1048575; // Maximum rows per sheet
                                int totalRows = dataTable.Rows.Count;
                                int sheetIndex = 1;
                                int currentRow = 0;
                                int u10ColumnIndex = 0;
                                bool alreadyHasu10ColumnIndex = false;

                                // Loop to handle multiple sheets
                                while (currentRow < totalRows)
                                {
                                    // Add a new worksheet
                                    var worksheet = workbook.Worksheets.Add($"Sheet{sheetIndex}");

                                    // Add the column headers
                                    for (int col = 0; col < dataTable.Columns.Count; col++)
                                    {
                                        worksheet.Cell(1, col + 1).Value = dataTable.Columns[col].ColumnName;
                                        if (!alreadyHasu10ColumnIndex && dataTable.Columns[col].ColumnName == "u10")
                                        {
                                            u10ColumnIndex = col;
                                            alreadyHasu10ColumnIndex = true;
                                        }
                                    }

                                    // Add the data rows for the current sheet
                                    int rowsInCurrentSheet = 0;
                                    while (rowsInCurrentSheet < maxRowsPerSheet && currentRow < totalRows)
                                    {
                                        for (int col = 0; col < dataTable.Columns.Count; col++)
                                        {
                                            // Get the cell value and convert it to string
                                            var cellValue = dataTable.Rows[currentRow][col];
                                            worksheet.Cell(rowsInCurrentSheet + 2, col + 1).Value = cellValue != DBNull.Value ? cellValue.ToString() : string.Empty;

                                            if (cellValue != DBNull.Value && col == u10ColumnIndex && Convert.ToDecimal(dataTable.Rows[currentRow][col]) < 0)
                                            {
                                                worksheet.Range(rowsInCurrentSheet + 2, 1, rowsInCurrentSheet + 2, dataTable.Columns.Count).Style.Fill.BackgroundColor = XLColor.Red;
                                            }
                                        }

                                        rowsInCurrentSheet++;
                                        currentRow++;
                                    }

                                    sheetIndex++;
                                }

                                // Save the workbook
                                workbook.SaveAs(filePath);
                                MessageBox.Show("Excel file has been saved successfully.", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"An error occurred while saving the Excel file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No data to export.", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        private void btnFilterDaily_Click(object sender, EventArgs e)
        {
            DataTable dataTable = (DataTable)dgImportedEra5.DataSource;
            if (dataTable != null)
            {
                if (!dataTable.Columns.Contains("date"))
                {
                    MessageBox.Show("'date' column does not exist.");
                    return;
                }

                var sortedRows = dataTable.AsEnumerable().OrderBy(row => row.Field<DateTime>("date")).CopyToDataTable();

                dgImportedEra5.DataSource = sortedRows;
            }
            else
            {
                MessageBox.Show("Filtered error.", "Filtered error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnFilterWeekly_Click(object sender, EventArgs e)
        {
            DataTable dataTable = (DataTable)dgImportedEra5.DataSource;
            if (dataTable != null)
            {
                if (!dataTable.Columns.Contains("date"))
                {
                    MessageBox.Show("'date' column does not exist.");
                    return;
                }

                var sortedRows = dataTable.AsEnumerable().OrderBy(row => row.Field<DateTime>("date")).CopyToDataTable();

                if (sortedRows != null)
                {
                    DataTable newTable = sortedRows.Clone(); // creating new table to not manipulate sortedRows itself
                    newTable.Columns.Add("week", typeof(int));

                    DateTime startDate = new DateTime(2023, 1, 1); // using this date as the basis for week counting

                    foreach (DataRow row in sortedRows.Rows)
                    {
                        DataRow newRow = newTable.NewRow(); // creates a new DataRow with the same schema as the table
                        newRow.ItemArray = row.ItemArray; // copies all values of a row to the newRow

                        DateTime currentDate = (DateTime)row["date"];

                        // (01/02/2023) - (01/01/2023) --> (1 / 7) + 1 = 0 + 1 = 1
                        // (01/08/2023) - (01/01/2023) --> (7 / 7) + 1 = 1 + 1 = 2
                        // (01/21/2023) - (01/01/2023) --> (20 / 7) + 1 = 2 + 1 = 3
                        // (01/22/2023) - (01/01/2023) --> (21 / 7) + 1 = 3 + 1 = 4
                        // ...
                        int daysDifference = (currentDate - startDate).Days;
                        int week = (daysDifference / 7) + 1;

                        newRow["week"] = week;

                        newTable.Rows.Add(newRow);
                    }

                    dgImportedEra5.DataSource = newTable;
                }
            }
            else
            {
                MessageBox.Show("Filtered error.", "Filtered error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
