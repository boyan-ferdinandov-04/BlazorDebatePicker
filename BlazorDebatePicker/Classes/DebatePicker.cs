using OfficeOpenXml;
using System.ComponentModel;

namespace BlazorDebatePicker.Classes
{
    public static class DebatePicker
    {
        static List<string> topics = null;
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

        public static List<string> ListTopic(string filePath)
        {
            topics ??= ReadTopics(filePath);
            List<string> randomTopics = new List<string>();
            Random random = new Random();
            int count = 4;


            for (int i = 0; i < count; i++)
            {
                int debateIndex = random.Next(topics.Count);
                randomTopics.Add(topics[debateIndex]);
                topics.RemoveAt(debateIndex);
            }
            return randomTopics;
        }
    }
}
