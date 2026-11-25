using System;
using System.Collections.Generic;
using static DevilWest.Symbol;
using static DevilWest.SlotInfo;
using static DevilWest.SlotSimulation;
using static DevilWest.SlotStats;
using System.Linq;
using OfficeOpenXml;
using System.IO;
using System.Reflection;
using DevilWest;


namespace DevilWest
{
    public class PaylinesCounter
    {
        public int GetPaylinesWins(Symbol[,] matrix, int[,] multiplierMatrix)
        {
            int paylinesTotalWin = 0;
            int symbolWin = 0;
            int substituteSymbolAmount = 0;
            Symbol winSymbol;
            foreach (var payline in Paylines.Values)
            {
                substituteSymbolAmount = 0;
                symbolWin = 0;
                winSymbol = matrix[payline[0], 0];
                if (winSymbol != Tomb)
                    symbolWin = GetSymbolWin(payline, matrix, winSymbol, out substituteSymbolAmount);
                if (symbolWin != 0)
                {
                    paylinesTotalWin += symbolWin * GetTotalMultiplierOfPayline(payline, multiplierMatrix, substituteSymbolAmount);
                    SymbolsHitsCount[winSymbol][substituteSymbolAmount - 1]++;
                }
            }
            return paylinesTotalWin;
        }

        public int GetPaylinesWinsWithShowingDetails(Symbol[,] matrix, int[,] multiplierMatrix)
        {
            int paylinesTotalWin = 0;
            int symbolWin = 0;
            int substituteSymbolAmount = 0;
            Symbol winSymbol;
            int i = 1;
            foreach (var payline in Paylines.Values)
            {
                substituteSymbolAmount = 0;
                symbolWin = 0;
                winSymbol = matrix[payline[0], 0];
                if (winSymbol != Tomb)
                    symbolWin = GetSymbolWin(payline, matrix, winSymbol, out substituteSymbolAmount);
                if (symbolWin != 0)
                {
                    int totalmultiplierOfPayline = GetTotalMultiplierOfPayline(payline, multiplierMatrix, substituteSymbolAmount);
                    paylinesTotalWin += symbolWin * totalmultiplierOfPayline;
                    Console.WriteLine($"Payline {i}, symbol {winSymbol}, amount {substituteSymbolAmount}, multiplier from wilds {totalmultiplierOfPayline}, win {symbolWin * totalmultiplierOfPayline} = ({symbolWin} * {totalmultiplierOfPayline})");
                    SymbolsHitsCount[winSymbol][substituteSymbolAmount - 1]++;
                }
                i++;
            }
            Console.WriteLine($"Total Win = {paylinesTotalWin}");
            return paylinesTotalWin;
        }

        private int GetSymbolWin(List<int> payline, Symbol[,] matrix, Symbol symbol, out int symbolAmount)
        {
            symbolAmount = 1;
            for (int i = 1; i < SlotWidth; i++)
            {
                if (matrix[payline[i], i] == Tomb || matrix[payline[i], i] == symbol)
                {
                    symbolAmount++;
                }
                else
                {
                    break;
                }
            }
            if (symbolAmount < 2)
                return 0;
            return PayTable[symbol][symbolAmount - 1];
        }

        private int GetTotalMultiplierOfPayline(List<int> payline, int[,] multiplierMatrix, int substituteSymbolsAmount)
        {
            int totalMultiplier = 1;
            int multiplier;
            for (int i = 0; i < substituteSymbolsAmount; i++)
            {
                multiplier = multiplierMatrix[payline[i], i];
                if (multiplier != 1)
                    totalMultiplier += multiplier;
            }
            if (totalMultiplier > 1)
                totalMultiplier--;
            return totalMultiplier;
        }

        public static void ShowMatrix(Symbol[,] matrix)
        {
            for (int i = 0; i < SlotHeight; i++)
            {
                for (int j = 0; j < SlotWidth; j++)
                {   
                    string symbolToWrite = matrix[i, j].ToString();
                    Console.Write($"{symbolToWrite}");
                    for (int k = 0; k < 12 - symbolToWrite.Length; k++)
                    {
                        Console.Write($" ");
                    }
                }
                Console.WriteLine();
            }
            Console.WriteLine();
        }
    }
}
