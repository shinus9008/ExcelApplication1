using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExcelApplication1
{
    public static class StrHelper
    {
        public static string RemoveEx(this string st, string text)
        {
            var index  = st.IndexOf(text);
            if (index >= 0)
            {
                return st.Remove(index, text.Length).Trim();
            }
            return st;
        }
    }

    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            var fileDialog = new FolderBrowserDialog();
            
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                Console.WriteLine("Open File:");
                Console.WriteLine(fileDialog.SelectedPath);

                var files = Directory.GetFiles(fileDialog.SelectedPath);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                foreach (var file in files)
                {
                    ExcelWork(file);
                }

                

                

               
                

                    

            }

            Console.WriteLine("Press key for exit");
            Console.ReadLine();
        }

        private static void ExcelWork(string file)
        {
            using (var excelPackage = new ExcelPackage(file))
            {
                var ws    = excelPackage.Workbook.Worksheets["ПЭ"];
                var start = ws.Dimension.Start.Row + 1;
                var end   = ws.Dimension.End.Row;

                for (int i = start; i <= end; i++)
                {
                    var cellC = ws.Cells[$"C{i}"];
                    if(!string.IsNullOrWhiteSpace(cellC.Text))
                    {
                        ws.Cells[$"F{i}"].Value = GetNewValue2(GetNewValue1(cellC.Text.Trim().ToUpper().Replace(',', '.')));
                    }

                    var cellG = ws.Cells[$"G{i}"];
                    if (!string.IsNullOrWhiteSpace(cellG.Text))
                    {
                        cellG.Value = cellG.Text
                            .ToUpper()
                            .RemoveEx("Ф.")
                            .RemoveEx("(")
                            .RemoveEx(")");            
                    }

                }

                excelPackage.Save();
            }
        }

       

        private static string GetNewValue1(string text)
        {
            foreach (var item in Excludes.OrderByDescending(x => x.Length))
            {
                var index = text.IndexOf(item);
                if (index == 0)               
                {
                    text = text.Remove(0, item.Length).Trim();
                    break;
                }
            }
           
            return text;
        }
        private static string GetNewValue2(string text)
        {


            foreach (var match in Matchs)
            {
                if (match.TryMath(ref text))
                    break;
            }

            return text;
        }


        private static IsMathe[] Matchs { get; } = new IsMathe[]
        {
            new Mathe_Key("136RVI-63", @"^136RVI[\-\s]*63"),

           
            
            new Mathe_Key("ADM483EAR"),

            new Mathe_Key("ATMEGA162-16AI"),

            new Mathe_Key("ATMEGA164A-AU"),
            new Mathe_Key("ATMEGA164A-AU", "ATMEGA164А-AU"),

            new Mathe_Key("BAV99"),

            //-----------------------------
            new Mathe_Key("HC-49/U3H-KX-3HT-12.000 МГЦ",  @"HC-49/U3H-KX-3HT[\-\s]+12.000\s*MГЦ"),

            new Mathe_Key("HC-49/U3H-KX-3HT-14.7456 МГЦ", @"HC-49/U3H-KX-3HT[\-\s]+14.7456\s*МГЦ"), //rus МГЦ
            new Mathe_Key("HC-49/U3H-KX-3HT-14.7456 МГЦ", @"HC-49/U3H-KX-3HT[\-\s]+14.7456\s*MГЦ"), //eng МГЦ

            new Mathe_Key("HC-49/US3H-KX-3HT-14.7456 МГЦ", @"HC-49/US3H-KX-3HT[\-\s]+14.7456\s*МГЦ"), //rus МГЦ
            new Mathe_Key("HC-49/US3H-KX-3HT-14.7456 МГЦ", @"HC-49/US3H-KX-3HT[\-\s]+14.7456\s*MГЦ"), //eng МГЦ

            new Mathe_Key("HC-49/US3H-KX-3HT-16 МГЦ", @"HC-49/US3H-KX-3HT[\-\s]+16\s*МГЦ"), //rus МГЦ
            new Mathe_Key("HC-49/US3H-KX-3HT-16 МГЦ", @"HC-49/US3H-KX-3HT[\-\s]+16\s*MГЦ"), //eng МГЦ

            new Mathe_Key("HC-49/US3H-KX-3HT 7372.8 КГЦ", @"HC-49/US3H-KX-3HT[\-\s]+7372.8\s*КГЦ"), //rus МГЦ
            new Mathe_Key("HC-49/US3H-KX-3HT 7372.8 КГЦ", @"HC-49/US3H-KX-3HT[\-\s]+7372.8\s*KГЦ"), //eng МГЦ

            
            //-----------------------------
            new Mathe_Key("RJ45"),
            new Mathe_Key("TAJB106M016R"),


            




            new Mathe_Key("TEN 3-4811", "TEN\\s+3\\-4811"),
            new Mathe_Key("PLD-10", "^PLD10"),
            new Mathe_Key("TEN 3-4811", "^TEN3-4811"),
            

            new Mathe_ddddKey("X7R"),
            new Mathe_ddddKey("NPO"),
            new Mathe_ddddKey("82"),

            new Mathe_ddddKey("0.125", "0\\.125"),
            new Mathe_ddddKey("0.25",  "0\\.25"),
            new Mathe_ddddKey("0.5",   "0\\.5"),

            new Mathe_ddddKey_Repl("Х7R", "X7R"), //-- Русский на английский
            new Mathe_ddddKey_Repl("NP0", "NPO"), //-- 0 на O
            
           
           

            
        };

        private static string[] Excludes { get; } = new string[]
        {
            "ВЫПРЯМИТЕЛЬНЫЙ ДИОД",
            "ВИЛКА",
            "ДИОД",
            "ДИОД ШОТТКИ",
            "ДИОДНАЯ СБОРКА",
            "ИНДИКАТОР",
            "ИНДУКТИВНОСТЬ",
            "КВАРЦЕВЫЙ РЕЗОНАТОР",
            "РАЗЪЕМ",
            "РЕЗИСТОР",
            "РЕЗОНАТОР",
            "СВЕТОДИОД",
            "ТРАНЗИСТОР",
            "ТРАНСФОРМАТОР",
            "ЧИП-ИНДУКТИВНОСТЬ",
            "ЧИП-КОНДЕНСАТОР",
            "ЧИП-ТАНТАЛ"
        };        
    }


    public interface IsMathe
    {
        bool TryMath(ref string text);
    }

    public class Mathe_Key : IsMathe
    {
        private readonly Regex regex;
        private readonly string key;

        public Mathe_Key(string key)
        {
            this.regex = new Regex(key);
            this.key = key;
        }
        public Mathe_Key(string key, string reg)
        {
            this.regex = new Regex(reg);
            this.key = key;
        }

        


        public bool TryMath(ref string text)
        {
            var m = regex.Match(text);
            if (m.Success)
            {
                text = key;
            }

            return m.Success;
        }
    }
    public class Mathe_ddddKey : IsMathe
    {
        private readonly Regex regex;
        private readonly string key;

        public Mathe_ddddKey(string key)
        {
            this.regex = new Regex(@"(?<DIGIT>\d\d\d\d)[\-\s]*" + key);
            this.key = key;
        }
        public Mathe_ddddKey(string key, string reg)
        {
            this.regex = new Regex("(?<DIGIT>\\d\\d\\d\\d)[\\-\\s]*" + reg);
            this.key = key;
        }


        public bool TryMath(ref string text)
        {
            var m = regex.Match(text);
            if (m.Success)
            {
                text = text = m.Groups["DIGIT"].Value + "-" + key;
            }

            return m.Success;
        }
    }

    public class Mathe_EXCLUDE : IsMathe
    {
        private readonly Regex regex;
        private readonly string key;

        public Mathe_EXCLUDE(string key)
        {
            this.regex = new Regex(key);
            this.key = key;
        }
        public Mathe_EXCLUDE(string key, string reg)
        {
            this.regex = new Regex(reg);
            this.key = key;
        }




        public bool TryMath(ref string text)
        {
            var m = regex.Match(text);
            if (m.Success)
            {
                text = key;
            }

            return m.Success;
        }
    }

    public class Mathe_ddddKey_Repl : IsMathe
    {
        private readonly Regex regex;
        private readonly string key;
    

        public Mathe_ddddKey_Repl(string key, string repl)
        {
            this.regex = new Regex(@"(?<DIGIT>\d\d\d\d)[\-\s]*" + key);
            this.key = repl;
           
        }
        public Mathe_ddddKey_Repl(string key, string reg, string repl)
        {
            this.regex = new Regex("(?<DIGIT>\\d\\d\\d\\d)[\\-\\s]*" + reg);
            this.key = key;         
        }


        public bool TryMath(ref string text)
        {
            var m = regex.Match(text);
            if (m.Success)
            {
                text = text = m.Groups["DIGIT"].Value + "-" + key;
            }

            return m.Success;
        }
    }
}
