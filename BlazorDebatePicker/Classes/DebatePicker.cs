using OfficeOpenXml;
using System.ComponentModel;

namespace BlazorDebatePicker.Classes
{
    public static class DebatePicker
    {
        public static List<string> ReadTopics(string filePath)
        {
            List<string> topics = new List<string>();
            FileInfo fileInfo = new FileInfo(filePath);
            
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int i = 2; i <= rowCount; i++)
                {
                    string topic = worksheet.Cells[i, 4].Value?.ToString();
                    if (!string.IsNullOrEmpty(topic))
                    {
                        topics.Add(topic);
                    }
                }
            }
            return topics;
        }

        public static string ListTopic(string filePath)
        {
            List<string> topics = ReadTopics(filePath);
            Random random = new Random();

            int debateIndex = random.Next(topics.Count);
            return topics[debateIndex];
        }
    }
}
