using SoftUniTableOfContentOOP.Interfaces;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System;
using System.IO;

namespace SoftUniTableOfContentOOP
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creating a document...
            WordDocument document = new WordDocument();
            document.EnsureMinimal();

            //Creating title...
            IAdd title = new MainClass();
            title.AddTitle(document, "Database Basics MS SQL Retake Exam - 10 Dec 2021");

            //Creating section...
            WSection mainSection = document.LastSection;

            //Adding paragraphs...
            IAdd para = new MainClass();
            para.AddParagraph(mainSection, para.AddHeading(1), "Section 1. DDL (30 pts)", "Random text for the first section");
            para.AddParagraph(mainSection, para.AddHeading(2), "1. Database design", "Design...");
            para.AddParagraph(mainSection, para.AddHeading(1), "Section 2. DML (10 pts)", "Random text for the second section");
            para.AddParagraph(mainSection, para.AddHeading(2), "2. Insert", "Inserting...");
            para.AddParagraph(mainSection, para.AddHeading(2), "3. Update", "Updating...");
            para.AddParagraph(mainSection, para.AddHeading(2), "4. Delete", "Deleting...");
            para.AddParagraph(mainSection, para.AddHeading(1), "Section 3. Querying (40 pts)", "Random text for the third section");
            para.AddParagraph(mainSection, para.AddHeading(2), "5. Aircraft", "Aircraft info...");
            para.AddParagraph(mainSection, para.AddHeading(3), "Example", "(empty)");
            para.AddParagraph(mainSection, para.AddHeading(2), "6. Pilots and Aircraft", "Pilots and Aircraft info...");
            para.AddParagraph(mainSection, para.AddHeading(3), "Example", "(empty)");
            para.AddParagraph(mainSection, para.AddHeading(2), "7. Top 20 Flight Destinations", "Selecting...");
            para.AddParagraph(mainSection, para.AddHeading(3), "Example", "(empty)");
            para.AddParagraph(mainSection, para.AddHeading(2), "8. Number of Flights for Each Aircraft", "Counting...");
            para.AddParagraph(mainSection, para.AddHeading(3), "Example", "(empty)");
            para.AddParagraph(mainSection, para.AddHeading(2), "9. Regular Passengers", "Passengers info...");
            para.AddParagraph(mainSection, para.AddHeading(3), "Example", "(empty)");
            para.AddParagraph(mainSection, para.AddHeading(2), "10. Full Info for Flight Destinations", "Destinations info...");
            para.AddParagraph(mainSection, para.AddHeading(3), "Example", "(empty)");
            para.AddParagraph(mainSection, para.AddHeading(1), "Section 4. Programmability (20 pts)", "Random text for the fourth section");
            para.AddParagraph(mainSection, para.AddHeading(2), "11. Find all Destinations by Email Address", "Finding...");
            para.AddParagraph(mainSection, para.AddHeading(3), "Example", "(empty)");
            para.AddParagraph(mainSection, para.AddHeading(2), "12. Full Info for Airports", "Info...");
            para.AddParagraph(mainSection, para.AddHeading(3), "Example", "(empty)");

            document.UpdateTableOfContents();

            //Creating the file on the PC...
            string path = @"D:\Documents\Projects\TableOfContentWord\TableOfContent.docx";
            Stream createFile = File.Create(Path.GetFullPath(path));
            document.Save(createFile, FormatType.Docx);
        }


    }
}
