using Parquet.Schema;
using Parquet;
using System.Data;
using OxyPlot;
using OxyPlot.Series;
using OxyPlot.WindowsForms;
using DataColumn = System.Data.DataColumn;
using OxyPlot.Axes;
using SharpCompress.Common;
using ClosedXML.Excel;
using System.Globalization;

namespace ArcVera_Tech_Test
{
    public partial class frmMain : Form
    {
        private DataTable dataTable;

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
                    dataTable = await ReadParquetFileAsync(filePath);
                    dgImportedEra5.DataSource = dataTable;
                    btnFilterDailyWeekly_SelectedOption(sender,e);
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

        private void PlotU10WeeklyValues(DataTable dataTable)
        {
            var plotModel = new PlotModel { Title = "Weekly u10 Values" };
            var lineSeries = new LineSeries { Title = "u10" };

            var groupedData = dataTable.AsEnumerable()
                .GroupBy(row => CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(
                    DateTime.Parse(row["date"].ToString()),
                    CalendarWeekRule.FirstDay,
                    DayOfWeek.Sunday))
                .Select(g => new
                {
                    //Group-by u10 values
                    WeekOfYear = g.Key,
                    U10Average = g.Average(row => Convert.ToDouble(row["u10"]))
                })
                .OrderBy(data => data.WeekOfYear);

            foreach (var data in groupedData)
            {
                lineSeries.Points.Add(new DataPoint(data.WeekOfYear, data.U10Average));
            }

            plotModel.Series.Add(lineSeries);
            plotView1.Model = plotModel;

        }

        public void btnFilterDailyWeekly_SelectedOption(object sender, EventArgs e)
        {
            ComboBox comboBox = null;

            if (sender is ComboBox)
            {
                comboBox = (ComboBox)sender;
            }
            else
            {
                comboBox = btnFilterDailyWeekly;
            }

            if (comboBox != null)
            {
                string selectedItem = (string)comboBox.SelectedItem;
                if (dataTable != null)
                {
                    if (selectedItem == "Daily")
                    {
                      
                        PlotU10DailyValues(dataTable);
                    }
                    else if (selectedItem == "Weekly")
                    {
                        PlotU10WeeklyValues(dataTable);
                    }
                }
            }
        }

        private void btnExportCsv_Click(object sender, EventArgs e)
        {
            //Check if exists table to export to csv
            if (dataTable != null)
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Csv files (*.csv)|*.csv|All files (*.*)|*.*";
                    saveFileDialog.Title = "Save as Csv File";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = saveFileDialog.FileName;

                        using (StreamWriter writer = new StreamWriter(filePath))
                        {
                            for (int i = 0; i < dataTable.Columns.Count; i++)
                            {
                                writer.Write(dataTable.Columns[i].ColumnName);
                                if (i < dataTable.Columns.Count - 1)
                                {
                                    writer.Write(",");
                                }
                            }
                            writer.WriteLine();

                            foreach (DataRow row in dataTable.Rows)
                            {
                                for (int i = 0; i < dataTable.Columns.Count; i++)
                                {
                                    writer.Write(row[i].ToString());
                                    if (i < dataTable.Columns.Count - 1)
                                    {
                                        writer.Write(",");
                                    }
                                }
                                writer.WriteLine();
                            }
                        }
                    }
                }
            }          
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            //Check if exists table to export to excel
            if (dataTable != null)
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    saveFileDialog.Title = "Save as Excel File";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = saveFileDialog.FileName;

                        int numberOfTables = (int)Math.Ceiling((double)dataTable.Rows.Count / 1000000);

                        for (int i = 0; i < numberOfTables; i++)
                        {
                            // Create a new DataTable with a subset of rows
                            DataTable smallerTable = dataTable.AsEnumerable().Skip(i * 1000000).Take(1000000).CopyToDataTable();
                            using var workbook = new XLWorkbook();
                            var worksheet = workbook.Worksheets.Add(smallerTable, "Worksheet");

                            // Apply conditional formatting to the entire row if the value in the 5th column is negative
                            var rowCount = worksheet.RowCount();
                            for (int row = 1; row <= rowCount; row++)
                            {
                                if (double.TryParse(worksheet.Cell(row, 5).Value.ToString(), out double value) && value < 0)
                                {
                                    worksheet.Row(row).Style.Fill.BackgroundColor = XLColor.Red;
                                }
                            }
                            string uniqueFilePath = Path.Combine(Path.GetDirectoryName(filePath),
                                Path.GetFileNameWithoutExtension(filePath) + (i + 1).ToString() + Path.GetExtension(filePath));
                            workbook.SaveAs(uniqueFilePath);
                        }
                    }
                }
            }
        }
    }
}
