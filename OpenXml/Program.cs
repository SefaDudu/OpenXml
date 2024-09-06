// See https://aka.ms/new-console-template for more information

using System.Reflection.Metadata;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Path = System.IO.Path;

string filePath = "RSU-BDDK.docx";
string testDate = "01/02/2024";
string companyName = "Test Company.";
string reportName = "Sızma Testi Örnek Rapor";
string reportStartDate = "01/03/2024";
string reportEndDate = "10/04/2024";

WordPlaceholderReplacer.ReplacePlaceholders(filePath, testDate, companyName, reportName, reportStartDate, reportEndDate);
Console.WriteLine("update");
public class WordPlaceholderReplacer
{
    public static void ReplacePlaceholders(string filePath, string testDate, string companyName, string reportName, string reportStartDate, string reportEndDate)
    {
        var directoryPath = Path.Combine(Directory.GetCurrentDirectory(), "file");
        var fullPath = Path.Combine(directoryPath, filePath);
        var newFilePath = Path.Combine(directoryPath, "new", "new_" + filePath);
        Directory.CreateDirectory(Path.Combine(directoryPath, "new"));
        File.Copy(fullPath, newFilePath, true);
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(newFilePath, true))
        {
            var docPart = wordDoc.MainDocumentPart;
            string docText = null;
            using (var reader = new StreamReader(docPart.GetStream()))
            {
                docText = reader.ReadToEnd();
            }
            var firstPageText = GetFirstPageText(docText);
            firstPageText = firstPageText.Replace("testDate", testDate);
            firstPageText = firstPageText.Replace("fullCompanyName", companyName);
            firstPageText = firstPageText.Replace("reportName", reportName);
            firstPageText = firstPageText.Replace("reportStartDate", reportStartDate);
            firstPageText = firstPageText.Replace("reportEndDate", reportEndDate);
            
            docText = docText.Replace(GetFirstPageText(docText), firstPageText);
            using (var writer = new StreamWriter(docPart.GetStream(FileMode.Create)))
            {
                writer.Write(docText);  
            }

            wordDoc.MainDocumentPart.Document.Save();
        }
    }
    private static string GetFirstPageText(string docText)
    {
        // Sayfa sonu karakterlerine göre bölme yap
        string[] pages = docText.Split(new string[] { "\f" }, StringSplitOptions.None);
    
        // İlk sayfayı döndür
        return pages[0];
    }
    }
