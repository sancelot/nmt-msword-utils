using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


namespace OpenXmlSdkTranslate
{
    class Program
    {
        static void Main(string[] args)
        {
            int text_id = 0;
            bool export_texts = false;
            bool translate = true;
            StreamWriter? textWriter = null;
            StreamReader? bluescore_file = null;
            var dictionary = new Dictionary<string, string>(); // dictionaire trad src => tgt language
            string input_filename = "input.docx";
            string output_filename = "output.docx";
            var arguments = new Dictionary<string, string>();
            for (int i = 0; i < args.Length; i += 2)
            {
                string key = args[i].Replace("-", "");
                string value = args[i + 1];
                arguments[key] = value;
            }

            // Now you can access arguments by name
            if (arguments.ContainsKey("extract_texts"))
            {
                Console.WriteLine("extracting texts");
                export_texts = true;
                translate = false;
            }
            if (arguments.ContainsKey("translate"))
            {
                Console.WriteLine("translating");
                export_texts = false;
                translate = true;
            }

            if (arguments.ContainsKey("input_filename"))
            {
                Console.WriteLine($"input_filename: {arguments["input_filename"]}");
                input_filename = arguments["input_filename"];
            }
            if (arguments.ContainsKey("output_filename"))
            {
                Console.WriteLine($"output_filename: {arguments["output_filename"]}");
                output_filename = arguments["output_filename"];
            }

            // Open the input.docx document.
            using (WordprocessingDocument inputDocument = WordprocessingDocument.Open(input_filename, true))
            {
                OpenXmlPackage doc = inputDocument.Clone(output_filename);
                doc.Dispose();
            }

            // create exported file
            if (export_texts)
            {
                textWriter = File.CreateText("input.txt");
            }


            if (translate)
            {
                // create a dictionaire : src -> tgt lang
                var inputLines = File.ReadLines("input_filtered.txt");
                var outputLines = File.ReadLines("output_filtered.txt");


                using (var e1 = inputLines.GetEnumerator())
                using (var e2 = outputLines.GetEnumerator())
                {
                    while (e1.MoveNext() && e2.MoveNext())
                    {
                        var inputLine = e1.Current;
                        var outputLine = e2.Current;

                        dictionary[inputLine] = outputLine;
                    }
                }

                // Print the dictionary to verify it's correct
                foreach (var kvp in dictionary)
                {
                    Console.WriteLine($"Key = {kvp.Key}, Value = {kvp.Value}");
                }

            }
            // Open the Word document.

            using (WordprocessingDocument document = WordprocessingDocument.Open(output_filename, true))
            {
                if (document.MainDocumentPart?.Document != null)
                {
                    // Loop through each paragraph in the document.
                    foreach (Paragraph paragraph in document.MainDocumentPart.Document.Descendants<Paragraph>())
                    {

                        // Concatenate the runs in the paragraph.
                        string line = ConcatenateRuns(paragraph);


                        // Translate all texts in the document.
                        //                    foreach (Text text in document.MainDocumentPart.Document.Descendants<Text>())
                        // ` returns all the `Text` elements in the document. You can then work with these `Text` objects directly.
                        //                    {
                        //                        Run? run = text?.Parent as Run;
                        //                        string? line = run?.InnerText;
                        if (!string.IsNullOrEmpty(line))// Process the line of text
                        {



                            string chaine = line;
                            string trimmed = String.Concat(line.Where(c => Char.IsLetter(c)));
                            if ((trimmed != "") && (chaine != ""))
                            {
                                // Console.WriteLine(line);

                                text_id++;

                                if (export_texts)
                                {
                                    textWriter?.WriteLine(line);
                                }

                                if (translate && dictionary.ContainsKey(line))
                                {

                                    string? translatedText = dictionary[line]; ;
                                    //string translatedText = TranslateText(line);
                                    if (translatedText != null)
                                    {
                                        // Remove all the runs in the paragraph.
                                        Run[] runs = paragraph.Elements<Run>().ToArray();
                                        paragraph.RemoveAllChildren<Run>();

                                        // Add a new run with the translated text.
                                        Run newRun = new Run(new Text(translatedText));

                                        // Clone the properties of the original run and apply them to the new run.
                                        if (runs.Length > 0)
                                        {
                                            RunProperties? runProps = runs[0].RunProperties;
                                            if (runProps != null)
                                            {
                                                newRun.PrependChild<RunProperties>((RunProperties)runProps.Clone());
                                            }
                                        }

                                        paragraph.AppendChild<Run>(newRun);
                                        // Replace the original text with the translated text.
                                        //text.Text = translatedText;

                                        // Get the parent Run element of the Text
                                        // Run? run = text.Ancestors<Run>().FirstOrDefault();
                                        // if (run != null)
                                        // {
                                        //     // change background color
                                        //     // Create a new RunProperties with the specified background color
                                        //     RunProperties runProperties = new RunProperties();

                                        //     Shading shading = new Shading() { Val = ShadingPatternValues.Solid, Color = "green", Fill = "00FF00" };
                                        //     runProperties.Append(shading);

                                        //     // Check if RunProperties already exist, if not, create one and add the shading
                                        //     RunProperties? existingRunProperties = run.Elements<RunProperties>().FirstOrDefault();

                                        //     if (existingRunProperties != null)
                                        //     {
                                        //         existingRunProperties.AppendChild(shading.CloneNode(true));
                                        //     }
                                        //     else
                                        //     {
                                        //         run.PrependChild<RunProperties>((RunProperties)runProperties.CloneNode(true));


                                        //     }
                                        // }
                                    }
                                }
                            }

                            // -----------------------
                        }
                    }
                }
                // Save the document.
                document.Save();

            }
            if (export_texts)
            {
                textWriter?.Close();

                // Filter duplicated_strings 
                string inputPath = "input.txt";
                string outputPath = "input_filtered.txt";

                HashSet<string> lines = new HashSet<string>(File.ReadLines(inputPath));
                File.WriteAllLines(outputPath, lines);
            }
        }
        private static string ConcatenateRuns(Paragraph paragraph)
        {
            StringBuilder sb = new StringBuilder();

            // Loop through each run in the paragraph.
            foreach (Run run in paragraph.Elements<Run>())
            {
                // Get the text element in the run.
                Text? text = run.Elements<Text>().FirstOrDefault();
                if (text != null)
                    // Append the text to the StringBuilder.
                    sb.Append(text.Text);
            }

            return sb.ToString();
        }
        private static string TranslateText(string text)
        {
            // Use a translation service to translate the text.
            // For example, you could use the Google Translate API.
            text = "123456" + text;
            return text;
        }
    }
}

