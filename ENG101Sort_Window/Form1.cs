using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;


namespace ENG101Sort_Window
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string filePath = ofd.FileName;
                Console.WriteLine(filePath);
                List<int> settings = GetSettings();
                GenerateFile(filePath, settings);
            }
            else
            {
                Console.WriteLine("Cancelled");
            }
            Application.Exit();

        }

        static List<int> GetSettings()
        {
            StreamReader sr = new StreamReader("settings.txt");
            List<int> settings = new List<int>();
            string line = sr.ReadLine();
            while (line != null)
            {
                if (line.Contains("="))
                {
                    string[] data = line.Split('=');
                    settings.Add(int.Parse(data[1])-1);
                }
                line = sr.ReadLine();
            }
            sr.Close();
            return settings;
        }

        static void GenerateFile(string filename, List<int> settings)
        {

            List<Student> students = new List<Student>();

            using (var reader = new StreamReader(filename))
            {
                // If more than one header line, uncomment the three code lines below, or even add more if more than two headers lines
                for (int i = 0; i < settings[0]+1; i++)
                {
                    var Qline = reader.ReadLine();
                    Console.WriteLine(Qline);
                }
                //var Qs = Qline.Split(',');
                //var headerline = reader.ReadLine();
                //Console.WriteLine(headerline);
                //var headers = headerline.Split(',');

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    Console.WriteLine(line);
                    var values = line.Split(',');
                    Student student = new Student();
                    student.InputFromCSVValues(values, settings);
                    students.Add(student);
                }
            }
            PDFCreation(students, filename + ".pdf");
            Console.WriteLine("Done");
            // Console.Read();

        }

        static void PDFCreation(List<Student> students, string filename)
        {
            var report = new PDFSortPages(students);
            var document = report.CreateDocument();
            var pdfRenderer = new PdfDocumentRenderer(true);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();
            pdfRenderer.Save(filename);
            Process.Start(filename);

        }


        class PDFSortPages
        {
            private Document document;
            private List<Student> students;

            public PDFSortPages(List<Student> studentData)
            {
                students = studentData;
            }

            public Document CreateDocument()
            {
                document = new Document();

                DefineStyles();

                foreach (Student student in students)
                {
                    Section section = document.AddSection();
                    section.PageSetup.TopMargin = "3.0cm";
                    section.PageSetup.LeftMargin = "1.5cm";
                    section.PageSetup.RightMargin = "1.5cm";

                    Table table = new Table();
                    table.Borders.Width = 0;
                    table.AddColumn(Unit.FromInch(5));
                    table.AddColumn(Unit.FromInch(2));

                    Row row = table.AddRow();
                    Cell cell = row.Cells[0];
                    Paragraph paragraph = cell.AddParagraph();
                    paragraph.Style = "NameStyle";
                    paragraph.AddFormattedText(student.FirstName + " " + student.LastName);
                    cell = row.Cells[1];
                    paragraph = cell.AddParagraph();
                    paragraph.Format.Alignment = ParagraphAlignment.Right;
                    paragraph.AddFormattedText(student.Gender);

                    row = table.AddRow();
                    cell = row.Cells[0];
                    paragraph = cell.AddParagraph();
                    paragraph.AddFormattedText(student.Email);
                    cell = row.Cells[1];
                    paragraph = cell.AddParagraph();
                    paragraph.Format.Alignment = ParagraphAlignment.Right;
                    paragraph.AddFormattedText(student.Ethnicity + "\n\n");

                    row = table.AddRow();
                    cell = row.Cells[0];
                    paragraph = cell.AddParagraph();
                    paragraph.AddFormattedText("");
                    cell = row.Cells[1];
                    paragraph = cell.AddParagraph();
                    paragraph.Format.Alignment = ParagraphAlignment.Right;
                    paragraph.AddFormattedText("Two Years in US HS:  " + student.USHS + "\n\n\n\n");


                    document.LastSection.Add(table);

                    table = new Table();
                    table.Borders.Width = 0;
                    table.AddColumn(Unit.FromInch(4.5));
                    table.AddColumn(Unit.FromInch(1.5));
                    table.AddColumn(Unit.FromInch(1.5));

                    row = table.AddRow();
                    cell = row.Cells[0];
                    paragraph = cell.AddParagraph();
                    paragraph.Format.Font.Bold = true;
                    paragraph.AddFormattedText("PROJECT PREFERENCES:\n\n");
                    cell = row.Cells[1];
                    paragraph = cell.AddParagraph();
                    paragraph.Format.Font.Bold = true;
                    paragraph.AddFormattedText("EXPERIENCE:\n\n");

                    string[] experienceTypes = new string[] { "Electronics", "Crafting", "Programming", "CAD", "Prototyping" };

                    for (int i = 1; i <= 5; i++)
                    {
                        row = table.AddRow();
                        cell = row.Cells[0];
                        paragraph = cell.AddParagraph();
                        paragraph.AddFormattedText("  " + i.ToString() + ".     " + student.Preference(i) + "\n\n");
                        cell = row.Cells[1];
                        paragraph = cell.AddParagraph();
                        paragraph.Format.Font.Bold = true;
                        paragraph.AddFormattedText(experienceTypes[i - 1] + "\n\n");
                        cell = row.Cells[2];
                        paragraph = cell.AddParagraph();
                        string expOutput;
                        switch (i)
                        {
                            case 1:
                                expOutput = student.Electronics;
                                break;
                            case 2:
                                expOutput = student.Crafting;
                                break;
                            case 3:
                                expOutput = student.Programming;
                                break;
                            case 4:
                                expOutput = student.CAD;
                                break;
                            case 5:
                                expOutput = student.Prototyping;
                                break;
                            default:
                                expOutput = "";
                                break;
                        }
                        paragraph.AddFormattedText(expOutput);
                    }

                    document.LastSection.Add(table);

                    table = new Table();
                    table.Borders.Width = 0;
                    table.AddColumn(Unit.FromInch(2));
                    table.AddColumn(Unit.FromInch(1));
                    row = table.AddRow();
                    cell = row.Cells[0];
                    paragraph = cell.AddParagraph();
                    paragraph.Format.Font.Bold = true;
                    paragraph.Format.Alignment = ParagraphAlignment.Right;
                    paragraph.AddFormattedText("\n\nAP Credits:  ");
                    cell = row.Cells[1];
                    paragraph = cell.AddParagraph();
                    paragraph.AddFormattedText("\n\n" + student.APCredits);

                    document.LastSection.Add(table);

                    /*table = new Table();
                    table.Borders.Width = 0;
                    table.AddColumn(Unit.FromInch(2));
                    table.AddColumn(Unit.FromInch(2));
                    row = table.AddRow();
                    cell = row.Cells[0];
                    paragraph = cell.AddParagraph();
                    paragraph.AddText("Time Zone:");
                    cell = row.Cells[1];
                    paragraph = cell.AddParagraph();
                    paragraph.AddText(student.TimeZone);

                    document.LastSection.Add(table);*/





                }

                return document;

            }

            private static void AddCell(Row row, int cellNumber, string inString)
            {
                Cell cell = row.Cells[cellNumber];
                Paragraph paragraph = cell.AddParagraph();
                paragraph.AddText(inString);
                return;
            }

            private void DefineStyles()
            {
                // Name
                Style style = document.Styles.AddStyle("NameStyle", "Normal");
                style.Font.Size = 18;
                style.Font.Bold = true;
                // style.Font.Name = "Calibri";

            }

        }

        class Student
        {
            public string FirstName;
            public string LastName;
            public string Email;
            public string Gender;
            public string Ethnicity;
            public string USHS;
            public string APCredits;
            public string Electronics;
            public string Crafting;
            public string Programming;
            public string CAD;
            public string Prototyping;
            public string Choice1;
            public string Choice2;
            public string Choice3;
            public string Choice4;
            public string Choice5;
            public string TimeZone;

            public void InputFromCSVValues(string[] values, List<int> settings)
            {
                FirstName = values[settings[1]];
                LastName = values[settings[2]];
                Email = values[settings[3]];
                Gender = values[settings[4]];
                Ethnicity = values[settings[5]];
                USHS = values[settings[6]];
                APCredits = values[settings[7]];
                Electronics = values[settings[8]];
                Crafting = values[settings[9]];
                Programming = values[settings[10]];
                CAD = values[settings[11]];
                Prototyping = values[settings[12]];
                if (values[settings[13]][0] == '"')
                {
                    Choice1 = values[settings[13]].Remove(0, 1);
                }
                else
                {
                    Choice1 = values[settings[13]];
                }

                Choice2 = values[settings[13]+1];
                Choice3 = values[settings[13]+2];
                Choice4 = values[settings[13]+3];
                Choice5 = values[settings[13]+4];
            }

            public string Preference(int choice)
            {
                switch (choice)
                {
                    case 1:
                        return this.Choice1;
                    case 2:
                        return this.Choice2;
                    case 3:
                        return this.Choice3;
                    case 4:
                        return this.Choice4;
                    case 5:
                        return this.Choice5;
                    default:
                        return "";
                }
            }
        }


    }

}
