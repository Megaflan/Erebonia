using System;

namespace Erebonia
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 3)
            {
                Console.WriteLine("Usage: Erebonia.exe {mode} {file}\n" +
                            "Modes: -e {file} {lang}: Export .XLSX file to a GetText Po file\n" +
                            "       -i {file} {original_xls}: Import a GetText Po file and outputs a .XLSX file\n" +
                            "\n" +
                            "Author: Megaflan\n" +
                            "Credits to Pleonex for the usage of Yarhl library and EPPlusSoftware for EPPlus\n" +
                            "(https://github.com/EPPlusSoftware/EPPlus/blob/develop/license.md)\n" +
                            "Special thanks to Darkmet98 for his Po2BinaryEasy solution and to Twn for CS3ScenarioDecompiler\n" +
                            "TraduSquare 2021");
                Console.ReadLine();
            }
            else
            {
                switch (args[0])
                {
                    case "-e":
                        XLS_Functions.Export(args[1], args[2]);
                        break;
                    case "-i":
                        XLS_Functions.Import(args[1], args[2]);
                        break;
                    default:
                        Console.WriteLine("Usage: Erebonia.exe {mode} {file}\n" +
                            "Modes: -e {file} {lang}: Export .XLSX file to a GetText Po file\n" +
                            "       -i {file} {original_xls}: Import a GetText Po file and outputs a .XLSX file\n" +
                            "\n" +
                            "Author: Megaflan\n" +
                            "Credits to Pleonex for the usage of Yarhl library and EPPlusSoftware for EPPlus\n" +
                            "(https://github.com/EPPlusSoftware/EPPlus/blob/develop/license.md)\n" +
                            "Special thanks to Darkmet98 for his Po2BinaryEasy solution and to Twn for CS3ScenarioDecompiler\n" +
                            "TraduSquare 2021");
                        Console.ReadLine();
                        break;
                }
            }            
        }
    }
}
