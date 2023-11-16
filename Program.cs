using ConsoleGPT;
using Microsoft.Office.Interop.Word;

namespace Docx_File_Maker_ChatGPT
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

            Console.WriteLine("Path of file:");
            List<string> fileText = new List<string>();
            object path = Console.ReadLine();
            object readOnly = false;
            object miss = System.Reflection.Missing.Value;
            if (path.ToString().EndsWith(".docx"))
            {
                Document docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
                foreach (Paragraph paragraph in docs.Paragraphs)
                {
                    fileText.Add(paragraph.Range.Text.ToString());
                }

                docs.Close();
            }
            else
            {
                fileText = File.ReadAllLines(path.ToString()).ToList();
            }
            Console.WriteLine("Output path: ");

            string outputPath = Console.ReadLine();

            Console.WriteLine("Website/Company name: ");

            string companyName = Console.ReadLine();

            Console.WriteLine("Website/Company purpose: ");

            string companyPurpose = Console.ReadLine();

            Console.WriteLine("Additional Requirements: ");

            string additionalRequirements = Console.ReadLine();

            List<System.Threading.Tasks.Task> tasks = new List<System.Threading.Tasks.Task>();

            for (int i = 0; i < fileText.Count(); i++)
            {
                if (String.IsNullOrEmpty(outputPath + @$"{fileText[i].Remove(fileText[i].Length - 1).Replace("?", "").Replace(":", " -").Replace("<", "").Replace(">", "").Replace("*", "").Replace('"', Convert.ToChar("'")).Replace("|", "").Trim()}") || String.IsNullOrWhiteSpace(outputPath + @$"{fileText[i].Remove(fileText[i].Length - 1).Replace("?", "").Replace(":", " -").Replace("<", "").Replace(">", "").Replace("*", "").Replace('"', Convert.ToChar("'")).Replace("|", "").Trim()}")) continue;

                //File.Create(outputPath + @$"{fileText[i].Remove(fileText[i].Length - 1).Replace("?", "").Replace(":", " -").Replace("<", "").Replace(">", "").Replace("*", "").Replace('"', Convert.ToChar("'")).Replace("|", "").Trim()}.docx");

                object filePath = outputPath + @$"{fileText[i].Remove(fileText[i].Length - 1).Replace("?", "").Replace(":", " -").Replace("<", "").Replace(">", "").Replace("*", "").Replace('"', Convert.ToChar("'")).Replace("|", "").Trim()}.docx";
                Document document = new Document();
                /*var chatGPTClient = new ChatGPTClient();

                string userMessage = Console.ReadLine();

                var chatResponse = await chatGPTClient.SendMessage(userMessage);

                List<string> response = new List<string>();

                foreach (var assistantMessage in chatResponse.Choices!.Select(c => c.Message))
                    response.Add(assistantMessage!.Content!.Trim());*/

                /*foreach (Microsoft.Office.Interop.Word.Range docRange in document.Words)
                    docRange.Text = "a\nb\nc";*/

                string pageName = @$"{fileText[i].Remove(fileText[i].Length - 1).Replace("?", "").Replace(":", " -").Replace("<", "").Replace(">", "").Replace("*", "").Replace('"', Convert.ToChar("'")).Replace("|", "").Trim()}";

                string userMessage = "Use UK grammar instead of USA grammar. Write content for the page 'PAGE_TITLE' for the company 'COMPANY_NAME' and also recommend they come to 'COMPANY_NAME' for COMPANY_PURPOSE in roughly 600 words. Don't start the content with a \"Welcome to COMPANY_NAME in LOCATION\" sort of sentence. The company is based in The UK so make it relevant but don't add unnecessary \"UK\" in the text.".Replace("PAGE_TITLE", pageName).Replace("COMPANY_NAME", companyName).Replace("COMPANY_PURPOSE", companyPurpose) + additionalRequirements;

                tasks.Add(GPTResponse(userMessage, document, filePath, miss));

                /*List<string> response = System.Threading.Tasks.Task.Run(() => GPTResponse(userMessage)).Result;

                document.Range().Text = String.Join("\n", response);

                if(File.Exists((string)filePath)) File.Delete((string)filePath);
                //Console.WriteLine(filePath);
                document.SaveAs2(ref filePath, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

                document.Close();*/
            }
            System.Threading.Tasks.Task.WaitAll(tasks.ToArray());
            Console.ReadKey();
        }

        static async System.Threading.Tasks.Task GPTResponse(string userMessage, Document document, object filePath, object miss)
        {
            var chatGPTClient = new ChatGPTClient();

            var chatResponse = await chatGPTClient.SendMessage(userMessage);

            List<string> response = new List<string>();

            if (chatResponse != null)
            {
                foreach (var assistantMessage in chatResponse.Choices!.Select(c => c.Message))
                    response.Add(assistantMessage!.Content!.Trim());
            }

            document.Range().Text = String.Join("\n", response);

            if (File.Exists((string)filePath)) File.Delete((string)filePath);
            //Console.WriteLine(filePath);
            document.SaveAs2(ref filePath, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

            Console.WriteLine(filePath);

            document.Close();
        }
    }
}