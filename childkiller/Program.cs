using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Windows.Forms;

class Program
{
    [STAThread]

    static void Main()
    {
        bool isRunning = true;

        while (isRunning)
        {
            Console.WriteLine("欢迎使用数学题目生成器！");
            Console.WriteLine("请选择操作：");
            Console.WriteLine("0. 退出程序");
            Console.WriteLine("1. 生成题目");

            if (int.TryParse(Console.ReadLine(), out int choice))
            {
                switch (choice)
                {
                    case 0:
                        isRunning = false;
                        Console.WriteLine("程序已退出。");
                        break;
                    case 1:
                        GenerateMathProblemsMenu();
                        break;
                    default:
                        Console.WriteLine("无效的选项，请重新输入。");
                        break;
                }
            }
            else
            {
                Console.WriteLine("请输入有效的数字选项。");
            }
        }
    }

    static void GenerateMathProblemsMenu()
    {
        List<string> mathProblems = new List<string>();

        while (true)
        {
            Console.WriteLine("请选择生成题目模式：");
            Console.WriteLine("1. 十以内加法");
            Console.WriteLine("2. 十以内减法");
            Console.WriteLine("3. 一百以内加法");
            Console.WriteLine("4. 一百以内减法");
            Console.WriteLine("5. 一千以内加法");
            Console.WriteLine("6. 一千以内减法");
            Console.WriteLine("0. 完成题目选择");

            if (int.TryParse(Console.ReadLine(), out int mode))
            {
                if (mode == 0)
                {
                    if (mathProblems.Count == 0)
                    {
                        Console.WriteLine("没有生成题目，请重新选择。");
                        continue;
                    }

                    Console.WriteLine("以下是生成的题目：");
                    PreviewMathProblems(mathProblems);

                    Console.WriteLine("将生成的题目保存到Word文档中吗？（Y/N）");
                    string saveToWord = Console.ReadLine();

                    if (saveToWord.ToLower() == "y")
                    {
                        SaveToWordDocument(mathProblems);
                        Console.WriteLine("题目已保存到指定位置。");
                    }

                    break;
                }

                Console.WriteLine("请输入题目数量：");

                if (int.TryParse(Console.ReadLine(), out int numberOfProblems))
                {
                    Console.WriteLine("允许负数吗？（Y/N）");
                    char allowNegative = Console.ReadLine().ToLower()[0];

                    if (allowNegative != 'y' && allowNegative != 'n')
                    {
                        Console.WriteLine("请输入有效的选项（Y/N）。");
                        continue;
                    }

                    if (mathProblems.Count + numberOfProblems > 10)
                    {
                        Console.WriteLine("题目总数超过10题，请重新选择题目数量或模式。");
                        continue;
                    }

                    mathProblems.AddRange(GenerateMathProblems(mode, numberOfProblems, allowNegative));
                }
                else
                {
                    Console.WriteLine("请输入有效的数字。");
                }
            }
            else
            {
                Console.WriteLine("请输入有效的数字选项。");
            }
        }
    }


    static List<string> GenerateMathProblems(int mode, int numberOfProblems, char allowNegative)
    {
        List<string> problems = new List<string>();
        Random random = new Random();

        for (int i = 0; i < numberOfProblems; i++)
        {
            string problem = "";

            if (mode == 1)
            {
                // 加法
                int operand1 = random.Next(1, 11);
                int operand2 = random.Next(1, 11);
                problem = $"{operand1} + {operand2}";
            }
            else if (mode == 2)
            {
                // 十以内的减法
                int operand1 = random.Next(1, 11);
                int operand2 = random.Next(1, 11);

                // 询问用户是否允许负数
                if (allowNegative == 'n' && operand1 < operand2)
                {
                    operand1 += operand2; // 交换操作数确保结果非负
                }

                problem = $"{operand1} - {operand2}";
            }
            else if (mode == 3)
            {
                // 一百以内的加法
                int operand1 = random.Next(1, 101);
                int operand2 = random.Next(1, 101);
                problem = $"{operand1} + {operand2}";
            }
            else if (mode == 4)
            {
                // 一百以内的减法
                int operand1 = random.Next(1, 101);
                int operand2 = random.Next(1, 101);

                // 询问用户是否允许负数
                if (allowNegative == 'n' && operand1 < operand2)
                {
                    operand1 += operand2; // 交换操作数确保结果非负
                }

                problem = $"{operand1} - {operand2}";
            }
            else if (mode == 5)
            {
                // 一千以内的加法
                int operand1 = random.Next(1, 1001);
                int operand2 = random.Next(1, 1001);
                problem = $"{operand1} + {operand2}";
            }
            else if (mode == 6)
            {
                // 一千以内的减法
                int operand1 = random.Next(1, 1001);
                int operand2 = random.Next(1, 1001);

                // 询问用户是否允许负数
                if (allowNegative == 'n' && operand1 < operand2)
                {
                    operand1 += operand2; // 交换操作数确保结果非负
                }

                problem = $"{operand1} - {operand2}";
            }
            // 其他情况返回空值

            problems.Add(problem);
        }

        return problems;
    }




    static void PreviewMathProblems(List<string> problems)
    {
        foreach (var problem in problems)
        {
            Console.WriteLine(problem);
        }
    }

    static void SaveToWordDocument(List<string> problems)
    {
        SaveFileDialog saveFileDialog = new SaveFileDialog();
        saveFileDialog.Filter = "Word文档 (*.docx)|*.docx";

        if (saveFileDialog.ShowDialog() == DialogResult.OK)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(saveFileDialog.FileName, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                foreach (var problem in problems)
                {
                    Paragraph paragraph = body.AppendChild(new Paragraph());
                    Run run = paragraph.AppendChild(new Run());
                    Text text = run.AppendChild(new Text(problem));
                }
            }
        }
    }
}
