using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Unicode;
using ExcelDataReader;

class Program
{
    static void Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); // Register the encoding provider

        string filePath = "C:/xxx/Documents/Projekte/SchulenNRW/CSV/GesamtePDF_korrekt.xlsx"; // Replace with the actual path to your Excel file
        string sheetName = "Münster"; // Replace with the name of the sheet you want to read

        List<string> columnValues = ReadColumnValues(filePath, sheetName, 1); // Assuming the third column index is 2 (zero-based)
        List<string> columnValues1 = ReadColumnValues(filePath, sheetName, 5); // Assuming the third column index is 2 (zero-based)
        List<string> columnValues2 = ReadColumnValues(filePath, sheetName, 10); // Assuming the third column index is 2 (zero-based)
        List<string> columnValues3 = ReadColumnValues(filePath, sheetName, 15); // Assuming the third column index is 2 (zero-based)
        List<string> columnValues4 = ReadColumnValues(filePath, sheetName, 19); // Assuming the third column index is 2 (zero-based)

        Random random = new Random();
        int randomIndex = random.Next(columnValues.Count);
        string randomValue = columnValues[randomIndex];
        int randomIndex1 = random.Next(columnValues1.Count);
        string randomValue1 = columnValues1[randomIndex1];
        int randomIndex2 = random.Next(columnValues2.Count);
        string randomValue2 = columnValues2[randomIndex2];
        int randomIndex3 = random.Next(columnValues3.Count);
        string randomValue3 = columnValues3[randomIndex3];
        int randomIndex4 = random.Next(columnValues4.Count);
        string randomValue4 = columnValues4[randomIndex4];


        columnValues.RemoveAt(randomIndex);
        columnValues1.RemoveAt(randomIndex1);
        columnValues2.RemoveAt(randomIndex2);
        columnValues3.RemoveAt(randomIndex3);
        columnValues4.RemoveAt(randomIndex4);

        int secondRandomIndex = random.Next(columnValues.Count);
        string secondRandomValue = columnValues[secondRandomIndex];
        int secondRandomIndex1 = random.Next(columnValues1.Count);
        string secondRandomValue1 = columnValues1[secondRandomIndex1];
        int secondRandomIndex2 = random.Next(columnValues2.Count);
        string secondRandomValue2 = columnValues2[secondRandomIndex2];
        int secondRandomIndex3 = random.Next(columnValues3.Count);
        string secondRandomValue3 = columnValues3[secondRandomIndex3];
        int secondRandomIndex4 = random.Next(columnValues4.Count);
        string secondRandomValue4 = columnValues4[secondRandomIndex4];


        // Store the random values in a text file
        string outputFilePath = "C:/xxx/Documents/Projekte/SchulenNRW/CSV/Münster_random.txt";
        using (StreamWriter writer = new StreamWriter(outputFilePath))
        {
            writer.WriteLine("Random value from Gesamtschule: " + randomValue);
            writer.WriteLine("Another random value from Gesamtschule (excluding the previous one): " + secondRandomValue);
            writer.WriteLine("Random value from Gymnasium: " + randomValue1);
            writer.WriteLine("Another random value from Gymnasium (excluding the previous one): " + secondRandomValue1);
            writer.WriteLine("Random value from Hauptschule: " + randomValue2);
            writer.WriteLine("Another random value from Hauptschule (excluding the previous one): " + secondRandomValue2);
            writer.WriteLine("Random value from Realschule: " + randomValue3);
            writer.WriteLine("Another random value from Realschule (excluding the previous one): " + secondRandomValue3);
            writer.WriteLine("Random value from Sekundarschule: " + randomValue4);
            writer.WriteLine("Another random value from Sekundarschule (excluding the previous one): " + secondRandomValue4);
        }
    }

    static List<string> ReadColumnValues(string filePath, string sheetName, int columnIndex)
    {
        List<string> columnValues = new List<string>();

        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                do
                {
                    if (reader.Name == sheetName)
                    {
                        while (reader.Read())
                        {
                            if (!reader.IsDBNull(columnIndex))
                            {
                                string value = reader.GetString(columnIndex);
                                columnValues.Add(value);
                            }
                        }
                        break;
                    }
                } while (reader.NextResult());
            }
        }

        return columnValues;
    }
}
