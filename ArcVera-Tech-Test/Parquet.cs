using System;
using System.IO;
using System.Linq;
using Parquet;
using Parquet.Data;
using System.Threading.Tasks;
using Parquet.Schema;
using Apache.Arrow;
using System.Data;
using static Microsoft.ML.DataViewSchema;
using static System.Net.Mime.MediaTypeNames;
using OfficeOpenXml;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using LargeXlsx;
using System.Data;

public class Class1
{
    static async Task Main(string[] args)
    {
        var options = new ParquetOptions { TreatByteArrayAsString = true };
        Stream fileStream = File.OpenRead("C:/Users/victor.nunes/OneDrive - Techbiz Forense Digital Ltda/Documents/Victor/ArcVeraC#_VictorKennedy/january-era5.parquet");
        var csvPath = "C:/Users/victor.nunes/OneDrive - Techbiz Forense Digital Ltda/Documents/Victor/ArcVeraC#_VictorKennedy/january-era5.csv";
        var excelPath = "C:/Users/victor.nunes/OneDrive - Techbiz Forense Digital Ltda/Documents/Victor/ArcVeraC#_VictorKennedy/january-era5.excel";
        var reader = await ParquetReader.CreateAsync(fileStream, options);

        var readerGroup = reader.OpenRowGroupReader(0);

        ParquetSchema check = reader.Schema;
        DataField[] dataFieldpandas = check.GetDataFields();

        // Create a new DataTable
        DataTable table = new DataTable();

        // Define columns and their types
        var columnDefinitions = new (string Name, Type Type)[]
        {
                ("longitude", typeof(float)),
                ("latitude", typeof(float)),
                ("date", typeof(DateTime)),
                ("time", typeof(TimeSpan)),
                ("u10", typeof(double))
        };

        // Adicione as colunas ao DataTable
        foreach (var columnDefinition in columnDefinitions)
        {
            table.Columns.Add(new System.Data.DataColumn(columnDefinition.Name, columnDefinition.Type));
        }

        // Leia os dados e adicione as linhas ao DataTable
        for (int i = dataFieldpandas.Length - 1; i >= 0; i--)
        {
            var columnData = await readerGroup.ReadColumnAsync(dataFieldpandas[i]);
            var dataList = columnData.Data.Cast<object>().ToList(); // Converta para uma lista
            for (int rowIndex = 0; rowIndex < dataList.Count; rowIndex++)
            {
                if (i == dataFieldpandas.Length - 1)
                {
                    // Adicione uma nova linha se for a primeira coluna
                    table.Rows.Add(table.NewRow());
                }

                // Se a coluna for "date" ou "time", extraia apenas a parte da data ou do horário
                if (columnDefinitions[i].Name == "date" && dataList[rowIndex] is DateTime date)
                {
                    table.Rows[rowIndex][columnDefinitions[i].Name] = date.Date;
                }
                else if (columnDefinitions[i].Name == "time" && dataList[rowIndex] is DateTime time)
                {
                    table.Rows[rowIndex][columnDefinitions[i].Name] = time.TimeOfDay;
                }
                else
                {
                    table.Rows[rowIndex][columnDefinitions[i].Name] = dataList[rowIndex];
                }
            }
        }

        ExportToCsv(table, csvPath);
        ExportToExcel(table, excelPath);

    }

    public static void ExportToCsv(DataTable table, string filePath)
    {
        using (StreamWriter writer = new StreamWriter(filePath))
        {
            // Write column names
            for (int i = 0; i < table.Columns.Count; i++)
            {
                writer.Write(table.Columns[i].ColumnName);
                if (i < table.Columns.Count - 1)
                {
                    writer.Write(",");
                }
            }
            writer.WriteLine();

            // Write rows
            foreach (DataRow row in table.Rows)
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    writer.Write(row[i].ToString());
                    if (i < table.Columns.Count - 1)
                    {
                        writer.Write(",");
                    }
                }
                writer.WriteLine();
            }
        }
    }

    public static void ExportToExcel(DataTable table, string filePath)
    {
        using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
        using (var xlsxWriter = new XlsxWriter(fileStream))
        {
            // Escreva os nomes das colunas
            xlsxWriter.BeginRow();
            for (int i = 0; i < table.Columns.Count; i++)
            {
                xlsxWriter.Write(table.Columns[i].ColumnName);
            }

            // Escreva os dados
            foreach (DataRow row in table.Rows)
            {
                xlsxWriter.BeginRow();
                for (int i = 0; i < row.ItemArray.Length; i++)
                {
                    xlsxWriter.Write(row.ItemArray[i]);
                }
            }
        }
    }
