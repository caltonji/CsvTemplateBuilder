using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.FileIO;
using System.IO;
using System.Text.RegularExpressions;


namespace ExcelReformattingApplication
{
    class Program
    {

        static string fillTemplate(string[] fields, string[] labels, string template)
        {

            string filledTemplate = String.Copy(template);
            Regex autoFillLabel = new Regex("{(.)+}");
            Console.WriteLine("chjec");
            while (autoFillLabel.IsMatch(filledTemplate))
            {
                // handle match here
                Console.WriteLine("hm,mm");
                Match match = Regex.Match(filledTemplate, "{(.)+}");
                string labelName = match.Value;
                labelName = labelName.Substring(1);
                labelName = labelName.Remove(labelName.Length - 1);
                Console.WriteLine("makes it here");
                int labelIndex = Array.IndexOf(labels, labelName);
                Console.WriteLine("and here?");
                if (labelIndex >= 0)
                {
                    string repl = fields[labelIndex];
                    string front = filledTemplate.Remove(match.Index);
                    string back = filledTemplate.Substring(match.Index + match.Value.Length);
                    filledTemplate = front + repl + back;
                    
                }
            }

            return filledTemplate;
        }


        static void Main(string[] args)
        {

            // Open all csv documents in the same folder as this
            string[] files = System.IO.Directory.GetFiles(Directory.GetCurrentDirectory(), "*.csv");
            string templatePath = Directory.GetCurrentDirectory() + "\\template.csv";
            Console.WriteLine(templatePath);
            int templateIndex = Array.IndexOf(files, templatePath);
            
            int successes = 0;
            List<string[]> failures = new List<string[]>();

            string template = String.Empty;
            string nameTemplate = String.Empty;

            if (templateIndex >= 0)
            {
                template = File.ReadAllText(files[templateIndex]);
                nameTemplate = template.Remove(template.IndexOf(','));

                var fileList = new List<string>(files);
                fileList.RemoveAt(templateIndex);
                files = fileList.ToArray();

                foreach (string path in files)
                {
                    //in your loop
                    //var newLine = string.Format("{0}", s);
                    //csv.AppendLine(newLine);
                    string folderPath = path.Substring(0, path.Length - 4);
                    System.IO.Directory.CreateDirectory(folderPath);
                    Console.WriteLine("created Directory");
                    string[] labels = new string[0];
                    using (TextFieldParser parser = new TextFieldParser(path))
                    {
                        Console.WriteLine("opened that shit");
                        parser.TextFieldType = FieldType.Delimited;
                        parser.SetDelimiters(",");
                        labels = parser.ReadFields();
                        Console.WriteLine("loaded labels: " + string.Join(",", labels));
                        while (!parser.EndOfData)
                        {
                            //Processing row
                            Console.WriteLine("entering parser");
                            string[] fields = parser.ReadFields();
                            Console.WriteLine("read files");

                            string newCsv = fillTemplate(fields, labels, template);
                            string fileName = fillTemplate(fields, labels, nameTemplate);
                            char[] invalidPathChars = Path.GetInvalidPathChars();
                            foreach (char c in invalidPathChars)
                            {
                                fileName = fileName.Replace(c.ToString(), String.Empty);
                            }
                            Console.WriteLine("entering try catch");
                            try {
                                int i = 0;
                                string tempFileName = String.Copy(fileName);
                                Console.WriteLine("checking folder " + folderPath + "/" + tempFileName);
                                while (File.Exists(folderPath + "/" + tempFileName + ".csv")) {
                                    i++;
                                    tempFileName = fileName + '_' + i.ToString();
                                }
                              
                                fileName = tempFileName;
                                Console.WriteLine("get here?");
                                File.WriteAllText(folderPath + "/" + fileName + ".csv", newCsv.ToString());
           
                                successes++;
                                Console.Write('.');
                            } catch(Exception e) {
                                failures.Add(fields);
                                Console.WriteLine("error " + e.Message);
                            }
                        }
                        
                    }
                    Console.WriteLine("Finished processing '" + Path.GetFileName(path) + "' with " + successes + " and " + failures.Count + " failures.");
                    if (failures.Count > 0)
                    {
                        Console.WriteLine("the failed rows are in the csv document labeled 'errors.csv'");

                        //before your loop
                        var errorCsv = new StringBuilder();
                        errorCsv.AppendLine(string.Join(",", labels));
                        foreach (string[] fields in failures)
                        {
                            errorCsv.AppendLine(string.Join(",", fields));
                        }
                        try
                        {
                            File.AppendAllText(Directory.GetCurrentDirectory() + "/errors.csv", errorCsv.ToString());
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Errors didn't save");
                        }
                    }
                    
                }


                
            }
            else
            {
                Console.WriteLine("No Template File was found, you must include a template file in this same folder as this .exe and it must be named 'template.csv'.  See the README for further instructions.");
            }
            Console.WriteLine("Press Enter to Exit");
            Console.ReadLine();
        }
    }
}
