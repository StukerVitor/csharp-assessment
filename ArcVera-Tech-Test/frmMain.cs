using Parquet.Schema;
using Parquet;
using System.Data;
using OxyPlot;
using OxyPlot.Series;
using OxyPlot.WindowsForms;
using DataColumn = System.Data.DataColumn;
using OxyPlot.Axes;
using System.Text;

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

                            // Add the column headers
                            IEnumerable<string> columnNames = dataTable.Columns.Cast<DataColumn>()
                                                     .Select(column => column.ColumnName);
                            csvContent.AppendLine(string.Join(",", columnNames));

                            // Add the rows
                            foreach (DataRow row in dataTable.Rows)
                            {
                                //Double single quotes to avoid confusion during csv export
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
            // Complete here
        }
    }
}
