﻿using System;
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
                    if (mathProblems.Count + numberOfProblems > 10)
                    {
                        Console.WriteLine("题目总数超过10题，请重新选择题目数量或模式。");
                        continue;
                    }

                    mathProblems.AddRange(GenerateMathProblems(mode, numberOfProblems));
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

    static List<string> GenerateMathProblems(int mode, int numberOfProblems)
    {
        List<string> problems = new List<string>();

        Random random = new Random();

        for (int i = 0; i < numberOfProblems; i++)
        {
            int operand1 = random.Next(1, 11);
            int operand2 = random.Next(1, 11);

            string problem = mode == 1 ? $"{operand1} + {operand2}" : $"{operand1} - {operand2}";

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
