using System;
using System.IO;

using Microsoft.Office.Interop.Word;

namespace ConsoleApp2
{
    class Program
    {
        static readonly string[] fileNames = new string[] { "Хохряков.docx", "Даниил 24.docx" };
        static Application app;
        static Document doc;
        static Paragraph prg;
        static Table table;
        static Microsoft.Office.Interop.Word.Range cell;
        static void CreateFiles(string[] names)
        {
            foreach (string name in names)
            {
                if (!new FileInfo(name).Exists)
                {
                    new FileInfo(name).Create().Close();
                }
            }
        }
        static void NewWordApp()
        {
            app ??= new() { Visible = true };
            doc = app.Documents.Add();
            prg = doc.Paragraphs.Add();
        }
        static void InsertText(string text, ref object style, string font, float size, short characterWidth, float firstLineIndent, WdParagraphAlignment alignment)
        {
            prg.Range.Text = text;
            prg.Range.set_Style(style);
            prg.Range.Font.Name = font;
            prg.Range.Font.Size = size;
            prg.IndentCharWidth(characterWidth);
            prg.Range.ParagraphFormat.FirstLineIndent = app.CentimetersToPoints(firstLineIndent);
            prg.Range.ParagraphFormat.Alignment = alignment;
            prg.Range.InsertParagraphAfter();
        }
        static void Save(string name)
        {
            doc.SaveAs(name);
            doc.Close(false);
        }
        static void NewTable()
        {
            table = doc.Tables.Add(prg.Range, 1, 4);
            table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            cell = table.Cell(1, 1).Range;
        }
        static double f(double x, double y, double z)
        {
            return 4 * Math.Pow(x, 2) + 3 * x - 2 * x * y + 4 * x * z;
        }
        static void Main()
        {
            CreateFiles(fileNames);
            NewWordApp();
            object styleName1 = "Заголовок", styleName2 = "Обычный", styleName3 = "Без интервала", styleName4 = "Цитата 2";
            InsertText("18.11.2021, Хохряков Даниил Андреевич, 220681, 1", ref styleName1, "Microsoft Himalaya", 20, 0, 0, WdParagraphAlignment.wdAlignParagraphCenter);
            InsertText("Свойство Range.Underline возвращает или задает тип подчеркивания, применённого к диапазону", ref styleName2, "Microsoft Himalaya", 18, 0, 1.5f, WdParagraphAlignment.wdAlignParagraphDistribute);
            prg.Range.InsertBreak();
            InsertText("Второй файл – создать таблицу значений заданной по варианту функции в выбранном пользователем диапазоне с заданным шагом. После таблицы вывести среднее значение функции на заданном интервале столько раз, каков Ваш номер в списке. Функция f(x,y,z)=4x^2+3x-2xy+4xz.", ref styleName3, "Microsoft Himalaya", 16, 0, 1, WdParagraphAlignment.wdAlignParagraphLeft);
            InsertText("\"Автостопом по галактике\", Дуглас Адамс. \"Как ни удивительно, единственной мыслью, посетившей вазу с петуниями, было: “Опять?! О нет, только не это”. Многие уверены, что, если бы мы знали, почему ваза с петуниями подумала именно так, мы могли бы понять природу вселенной гораздо лучше, чем понимаем сейчас.\"", ref styleName4, "Microsoft Himalaya", 14, 0, 0, WdParagraphAlignment.wdAlignParagraphRight);
            Save(fileNames[0]);
            Console.Write("Начало диапазона по x: ");
            double ax = Convert.ToDouble(Console.ReadLine());
            Console.Write("Конец диапазона по x: ");
            double bx = Convert.ToDouble(Console.ReadLine());
            Console.Write("Шаг по x: ");
            double hx = Convert.ToDouble(Console.ReadLine());
            Console.Write("Начало диапазона по y: ");
            double ay = Convert.ToDouble(Console.ReadLine());
            Console.Write("Конец диапазона по y: ");
            double by = Convert.ToDouble(Console.ReadLine());
            Console.Write("Шаг по y: ");
            double hy = Convert.ToDouble(Console.ReadLine());
            Console.Write("Начало диапазона по z: ");
            double az = Convert.ToDouble(Console.ReadLine());
            Console.Write("Конец диапазона по z: ");
            double bz = Convert.ToDouble(Console.ReadLine());
            Console.Write("Шаг по z: ");
            double hz = Convert.ToDouble(Console.ReadLine());
            NewWordApp();
            NewTable();
            cell.Text = "x";
            cell = cell.Next();
            cell.Text = "y";
            cell = cell.Next();
            cell.Text = "z";
            cell = cell.Next();
            cell.Text = "f";
            double avg = 0;
            for (double x = ax; x <= bx; x += hx)
            {
                for (double y = ay; y <= by; y += hy)
                {
                    for (double z = az; z <= bz; z += hz)
                    {
                        cell = table.Rows.Add().Cells[1].Range;
                        cell.Text = x.ToString();
                        cell = cell.Next();
                        cell.Text = y.ToString();
                        cell = cell.Next();
                        cell.Text = z.ToString();
                        cell = cell.Next();
                        cell.Text = f(x, y, z).ToString();
                        avg += f(x, y, z);
                    }
                }
            }
            avg /= table.Rows.Count - 1;
            for (int i = 0; i < 24; ++i)
            {
                prg.Range.Text = $"Среднее значение функции = {avg}\n";
            }
            Save(fileNames[1]);
            app.Quit();
        }
    }
}
