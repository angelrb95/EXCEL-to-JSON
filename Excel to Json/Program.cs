using System;
using System.Data;
using System.IO;
using ExcelDataReader;
using Newtonsoft.Json;

class Program
{
    static void Main()
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Ruta al archivo Excel
        string excelFilePath = "C:\\Users\\a.rodriguez\\Downloads\\listadopet.xls";

        // Leer el archivo Excel y convertirlo a DataSet
        DataSet dataSet;
        using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var config = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true // Usa la primera fila como nombres de columna
                    }
                };
                dataSet = reader.AsDataSet(config);
            }
        }

        // Obtener la tabla de datos
        DataTable dataTable = dataSet.Tables[0];
        int totalRows = dataTable.Rows.Count;
        int rowsPerFile = totalRows / 3;

        // Dividir y guardar cada parte en archivos JSON separados
        for (int i = 0; i < 3; i++)
        {
            DataTable partTable = dataTable.Clone(); // Clonar la estructura de la tabla
            int startRow = i * rowsPerFile;
            int endRow = (i == 2) ? totalRows : startRow + rowsPerFile; // Manejar la última parte

            for (int j = startRow; j < endRow; j++)
            {
                partTable.ImportRow(dataTable.Rows[j]);
            }

            // Convertir la parte del DataTable a JSON
            var jsonResult = JsonConvert.SerializeObject(partTable, Newtonsoft.Json.Formatting.Indented);

            // Guardar el JSON en un archivo
            string jsonFilePath = $"C:\\Users\\a.rodriguez\\Downloads\\listadopet_part{i + 1}.json";
            File.WriteAllText(jsonFilePath, jsonResult);
        }

        Console.WriteLine("El archivo Excel ha sido convertido y dividido en 3 archivos JSON exitosamente.");
    }
}
