using System;
using System.Collections.Generic;
using static DevilWest.Symbol;
using static DevilWest.SlotInfo;
using static DevilWest.SlotStats;
using System.Linq;
using System.IO;
using OfficeOpenXml;
using System.Reflection;
using System.Collections.Concurrent;

namespace DevilWest
{
    public enum Bins
    {
        Bin1,
        Bin2, 
        Bin3,
    }

    static class SlotInfo
    {
        public static double CostToPlay = 10;
        public const int SlotWidth = 5;
        public const int SlotHeight = 4;
        public const int MAXWIN = 50000;
        public static int Config = 1;

        public static List<Dictionary<int, List<Symbol>>> BGReelsSets = new List<Dictionary<int, List<Symbol>>>();
        public static List<Dictionary<int, List<Symbol>>> FSReelsSets = new List<Dictionary<int, List<Symbol>>>();
        public static Dictionary<int, List<Symbol>> ReelSetOfBuyFeature = new Dictionary<int, List<Symbol>>();
        public static Dictionary<Symbol, int[]> PayTable = new Dictionary<Symbol, int[]>();
        public static Dictionary<int, List<int>> Paylines = new Dictionary<int, List<int>>();
        public static Dictionary<int, int> WildsMultipliersOfBG = new Dictionary<int, int>();
        public static Dictionary<int, int> WildsMultipliersOfFS = new Dictionary<int, int>();
        public static Dictionary<int, int> ChosenSymbolsMultipliers = new Dictionary<int, int>();
        public static Dictionary<Dictionary<int, List<Symbol>>, int> FSReelsWeights = new Dictionary<Dictionary<int, List<Symbol>>, int>();
        public static Dictionary<Dictionary<int, List<Symbol>>, int> BGReelsWeights = new Dictionary<Dictionary<int, List<Symbol>>, int>();
        public static Dictionary<Dictionary<int, List<Symbol>>, Dictionary<Symbol, int>> InnerReelSetsBG = new Dictionary<Dictionary<int, List<Symbol>>, Dictionary<Symbol, int>>();
        public static Dictionary<Dictionary<int, List<Symbol>>, Dictionary<Symbol, int>> InnerReelSetsFS = new Dictionary<Dictionary<int, List<Symbol>>, Dictionary<Symbol, int>>();
        public static Dictionary<int, int> AdditionalFSAmountsTable = new Dictionary<int, int>();
        public static Dictionary<int, int> StartFreeSpinsAmount = new Dictionary<int, int>();
        public static Dictionary<Symbol, int> WeightsToChooseSymbol = new Dictionary<Symbol, int>();


        public static void ImportInfo()
        {

            using (var package = new ExcelPackage(new FileInfo(@"../../DevilWestMath.xlsx")))
            {
                try
                {
                    var workbook = package.Workbook;
                    ImportPaylines(workbook.Worksheets["Paylines"]);
                    ImportPaytable(workbook.Worksheets["Paytable"]);
                    ImportBaseSpinsWeights(workbook.Worksheets[$"BG Reels Weights {Config}"]);
                    CostToPlay = Convert.ToDouble(workbook.Worksheets["Summary"].Cells[$"C{2 + Config}"].Value);
                    for (int i = 0; i < BGReelsWeights.Count; i++)
                    {
                        ImportReelsFromSheetTo(workbook.Worksheets[$"Base Game Reels {i + 1}"], BGReelsSets[i]);
                        ImportInnerSymbolsOfReelSetBG(workbook.Worksheets[$"Base Game Reels {i + 1}"], BGReelsSets[i]);
                    }
                    ImportFreeSpinsWeights(workbook.Worksheets["Weights of FS Reels"]);
                    for (int i = 0; i < FSReelsWeights.Count; i++)
                    {
                        ImportReelsFromSheetTo(workbook.Worksheets[$"Free Spins Reels {i + 1}"], FSReelsSets[i]);
                        ImportInnerSymbolsOfReelSetFS(workbook.Worksheets[$"Free Spins Reels {i + 1}"], FSReelsSets[i]);
                    }
                    ImportSymbolsMultipliers(workbook.Worksheets[$"Symbols Multipliers"]);
                    ImportFSAmountsTable(workbook.Worksheets["Free Spins Amounts Table"]);
                    ImportWeightsToChooseActivatedSymbol(workbook.Worksheets["Weights of Choosing Symbol"]);
                    ImportReelsFromSheetTo(workbook.Worksheets["Reel Set Of Buy Feature"], ReelSetOfBuyFeature);
                }
                finally
                {
                    package.Dispose();
                }
            }
        }

        private static void ImportWeightsToChooseActivatedSymbol(ExcelWorksheet worksheet)
        {
            for (int i = 3; worksheet.Cells[$"B{i}"].Value != null; i++)
            {
                WeightsToChooseSymbol.Add(ConvertCellToSymbol(worksheet.Cells[$"B{i}"].Value), Convert.ToInt32(worksheet.Cells[$"C{i}"].Value));
            }
        }

        private static void ImportFSAmountsTable(ExcelWorksheet worksheet)
        {
            for (int i = 3; worksheet.Cells[$"F{i}"].Value != null; i++)
            {
                AdditionalFSAmountsTable.Add(Convert.ToInt32(worksheet.Cells[$"F{i}"].Value), Convert.ToInt32(worksheet.Cells[$"G{i}"].Value));
            }
            for (int i = 3; worksheet.Cells[$"B{i}"].Value != null; i++)
            {
                StartFreeSpinsAmount.Add(Convert.ToInt32(worksheet.Cells[$"B{i}"].Value), Convert.ToInt32(worksheet.Cells[$"C{i}"].Value));
            }
        }

        private static void ImportBaseSpinsWeights(ExcelWorksheet worksheet)
        {
            for (int i = 3; worksheet.Cells[$"D{i}"].Value != null; i++)
            {
                if (BGReelsSets.Count <= i - 3)
                    BGReelsSets.Add(new Dictionary<int, List<Symbol>>());
                BGReelsWeights.Add(BGReelsSets[i - 3], Convert.ToInt32(worksheet.Cells[$"D{i}"].Value));
            }
        }

        private static void ImportFreeSpinsWeights(ExcelWorksheet worksheet)
        {
            for (int i = 3; worksheet.Cells[$"D{i}"].Value != null ; i++)
            {
                if (FSReelsSets.Count <= i - 3)
                    FSReelsSets.Add(new Dictionary<int, List<Symbol>>());
                FSReelsWeights.Add(FSReelsSets[i - 3], Convert.ToInt32(worksheet.Cells[$"D{i}"].Value));
            }
        }

        private static void ImportSymbolsMultipliers(ExcelWorksheet worksheet)
        {
            for (int i = 4; worksheet.Cells[$"B{i}"].Value != null; i++)
            {
                WildsMultipliersOfBG.Add(Convert.ToInt32(worksheet.Cells[$"B{i}"].Value), Convert.ToInt32(worksheet.Cells[$"C{i}"].Value));
                WildsMultipliersOfFS.Add(Convert.ToInt32(worksheet.Cells[$"G{i}"].Value), Convert.ToInt32(worksheet.Cells[$"H{i}"].Value));
                ChosenSymbolsMultipliers.Add(Convert.ToInt32(worksheet.Cells[$"L{i}"].Value), Convert.ToInt32(worksheet.Cells[$"M{i}"].Value));
            }
        }

        private static void ImportPaylines(ExcelWorksheet worksheet)
        {
            for (int i = 9; worksheet.Cells[i, 1].Value != null; i++)
            {
                int paylineNumber = Convert.ToInt32(worksheet.Cells[i, 1].Value);
                Paylines.Add(paylineNumber, new List<int>());
                for (int j = 2; worksheet.Cells[i, j].Value != null; j++)
                {
                    int paylineDigit = Convert.ToInt32(worksheet.Cells[i, j].Value);
                    Paylines[paylineNumber].Add(paylineDigit);
                }
            }
        }

        private static void ImportReelsFromSheetTo(ExcelWorksheet worksheet, Dictionary<int, List<Symbol>> reelSetForImportInto)
        {
            int i;
            for (i = 0; i < SlotWidth; i++)
            {
                reelSetForImportInto.Add(i, new List<Symbol>());
                for (int j = 4; worksheet.Cells[j, i + 3].Value != null; j++)
                {
                    reelSetForImportInto[i].Add(ConvertCellToSymbol(worksheet.Cells[j, i + 3].Value));
                }
            }
        }
        

        private static void ImportPaytable(ExcelWorksheet worksheet)
        {
            for (int i = 4; worksheet.Cells[i, 2].Value != null; i++)
            {
                Symbol symbol = ConvertCellToSymbol(worksheet.Cells[i, 2].Value);
                PayTable.Add(symbol, new int[SlotWidth]);
                for (int j = 0; worksheet.Cells[i, j + 3].Value != null; j++)
                {
                    PayTable[symbol][j] = Convert.ToInt32(worksheet.Cells[i, j + 3].Value);
                }
            }
        }

        private static void ImportInnerSymbolsOfReelSetBG(ExcelWorksheet worksheet, Dictionary<int, List<Symbol>> reelSet)
        {
            if (InnerReelSetsBG.ContainsKey(reelSet) == false)
                InnerReelSetsBG.Add(reelSet, new Dictionary<Symbol, int>());
            for (int i = 4; worksheet.Cells[$"H{i}"].Value != null; i++)
            {
                Symbol symbolToAdd = ConvertCellToSymbol(worksheet.Cells[$"H{i}"].Value);
                if (InnerReelSetsBG[reelSet].ContainsKey(symbolToAdd) == false)
                {
                    InnerReelSetsBG[reelSet].Add(symbolToAdd, 1);
                }
                else
                {
                    InnerReelSetsBG[reelSet][symbolToAdd]++;
                }
            }
        }

        private static void ImportInnerSymbolsOfReelSetFS(ExcelWorksheet worksheet, Dictionary<int, List<Symbol>> reelSet)
        {
            if (InnerReelSetsFS.ContainsKey(reelSet) == false)
                InnerReelSetsFS.Add(reelSet, new Dictionary<Symbol, int>());
            for (int i = 4; worksheet.Cells[$"H{i}"].Value != null; i++)
            {
                Symbol symbolToAdd = ConvertCellToSymbol(worksheet.Cells[$"H{i}"].Value);
                if (InnerReelSetsFS[reelSet].ContainsKey(symbolToAdd) == false)
                {
                    InnerReelSetsFS[reelSet].Add(symbolToAdd, 1);
                }
                else
                {
                    InnerReelSetsFS[reelSet][symbolToAdd]++;
                }
            }
        }

        public static Symbol ConvertCellToSymbol(object cellValue)
        {
            string cellsString = cellValue.ToString();
            foreach (Symbol symbol in Enum.GetValues(typeof(Symbol)))
            {
                if (cellsString == symbol.ToString())
                {
                    return symbol;
                }
            }
            throw new Exception("ConvertCellToSymbol");
        }
    }
}
