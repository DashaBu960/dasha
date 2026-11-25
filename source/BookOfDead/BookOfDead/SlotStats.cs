using System;
using System.Collections.Generic;
using static DevilWest.Symbol;
using static DevilWest.SlotInfo;
using static DevilWest.SlotStats;
using System.Linq;
using System.Threading;
using OfficeOpenXml;
using System.IO;
using System.Reflection;
using System.Xml;
using System.Net;
using System.Collections.Concurrent;

namespace DevilWest
{
    static class SlotStats
    {
        public const long IterationsCount = (long)3E7;  // стояло 1E9
        public static long CurrentIteration = 0;
        public static ConcurrentDictionary<Symbol, long[]> SymbolsHitsCount = new ConcurrentDictionary<Symbol, long[]>();

        public static ConcurrentDictionary<int, long> WinsPerTotalSpin = new ConcurrentDictionary<int, long>();
        public static ConcurrentDictionary<int, long> WinsPerBGOnlySpin = new ConcurrentDictionary<int, long>();
        public static ConcurrentDictionary<int, long> WinsPerFSRound = new ConcurrentDictionary<int, long>();
        public static List<int> Intervals = new List<int> { -1, 0, 1, 2, 3, 5, 10, 15, 20, 30, 50, 75, 100, 150, 200, 250, 300, 400, 500, 1000, 1500, 2500, 5000 };
        public static ConcurrentDictionary<int, long> IntervalTotalWinsX = new ConcurrentDictionary<int, long>();
        public static ConcurrentDictionary<int, long> IntervalBaseGameWinsX = new ConcurrentDictionary<int, long>();
        public static ConcurrentDictionary<int, long> IntervalFSRoundWinsX = new ConcurrentDictionary<int, long>();
        public static ConcurrentDictionary<Symbol, int> SymbolsChosenCount = new ConcurrentDictionary<Symbol, int>();
        public static ConcurrentDictionary<int, long> WinsPerFreeSpin = new ConcurrentDictionary<int, long>();
        public static ConcurrentDictionary<int, long> FreeSpinsCountPerRound = new ConcurrentDictionary<int, long>();
        public static ConcurrentDictionary<int, long> WildsAmountInTheEndOfFSRound = new ConcurrentDictionary<int, long>();

        public static ConcurrentDictionary<int, long> ROE = new ConcurrentDictionary<int, long>();
        public static ConcurrentDictionary<Dictionary<int, List<Symbol>>, ConcurrentDictionary<int, long>> WinsFromReelSetsBGOnly = new ConcurrentDictionary<Dictionary<int, List<Symbol>>, ConcurrentDictionary<int, long>>();
        public static long WinRateOfLuckyBet = 1;
        public static long WinRate = 1;
        public static long WinRateOf10x = 1;
        public static long WinRateOf100x = 1;
        public static long WinRateOf1000x = 1;
        public static long HitRate = 1;

        public static long FSTriggersCount = 0;   // bil 1
        public static long FSCount = 0;

        public static double StdDev = 0;
        public static double TotalRTP = 0;
        public static double BGOnlyRTP = 0;
        public static double ScattersRTP = 0;
        public static double FreeSpinsRTP = 0;
        public static double ConfidenceInterval = 0;
        public static int MaxWin = 0;

        public static void ResetStats()
        {
            CurrentIteration = 0;
            SymbolsHitsCount = new ConcurrentDictionary<Symbol, long[]>();

            WinsPerTotalSpin = new ConcurrentDictionary<int, long>();
            WinsPerBGOnlySpin = new ConcurrentDictionary<int, long>();
            WinsPerFSRound = new ConcurrentDictionary<int, long>();
            Intervals = new List<int> { -1, 0, 1, 2, 3, 5, 10, 15, 20, 30, 50, 75, 100, 150, 200, 250, 300, 400, 500, 1000, 1500, 2500, 4999 };
            IntervalTotalWinsX = new ConcurrentDictionary<int, long>();
            IntervalBaseGameWinsX = new ConcurrentDictionary<int, long>();
            IntervalFSRoundWinsX = new ConcurrentDictionary<int, long>();
            SymbolsChosenCount = new ConcurrentDictionary<Symbol, int>();
            WildsAmountInTheEndOfFSRound = new ConcurrentDictionary<int, long>();

            WinsPerFreeSpin = new ConcurrentDictionary<int, long>();
            FreeSpinsCountPerRound = new ConcurrentDictionary<int, long>();

            ROE = new ConcurrentDictionary<int, long>();
            WinsFromReelSetsBGOnly = new ConcurrentDictionary<Dictionary<int, List<Symbol>>, ConcurrentDictionary<int, long>>();
            WinRateOfLuckyBet = 1;
            WinRate = 1;
            WinRateOf10x = 1;
            WinRateOf100x = 1;
            WinRateOf1000x = 1;
            HitRate = 1;

            FSTriggersCount = 1;
            FSCount = 0;

            StdDev = 0;
            TotalRTP = 0;
            BGOnlyRTP = 0;
            ScattersRTP = 0;
            FreeSpinsRTP = 0;
            ConfidenceInterval = 0;
            MaxWin = 0;


            for (int i = 0; i < BGReelsSets.Count; i++)
            {
                WinsFromReelSetsBGOnly.TryAdd(BGReelsSets[i], new ConcurrentDictionary<int, long>());
            }
            foreach (Symbol item in Enum.GetValues(typeof(Symbol)))
            {
                if (item != Symbol.Blank && item != Symbol.Inner && item != Tomb)
                {
                    SymbolsHitsCount[item] = new long[SlotWidth];
                    for (int i = 0; i < SlotWidth; i++)
                    {
                        SymbolsHitsCount[item][i] = 0;
                    }
                }
            }
            foreach (var interval in Intervals)
            {
                IntervalTotalWinsX.TryAdd(interval, 0);
                IntervalBaseGameWinsX.TryAdd(interval, 0);
                IntervalFSRoundWinsX.TryAdd(interval, 0);
            }
        }

        public static void AddWinTo(int win, ConcurrentDictionary<int, long> winsCountToAdd)
        {
            winsCountToAdd.AddOrUpdate(win, 1, (key, oldValue) => oldValue + 1);
        }

        public static void AddWinTo(int win, ConcurrentDictionary<int, int> winsCountToAdd)
        {
            winsCountToAdd.AddOrUpdate(win, 1, (key, oldValue) => oldValue + 1);
        }

        public static void AddWinTo<T>(int key, T value, ConcurrentDictionary<int, List<T>> winsCountToAdd)
        {
            winsCountToAdd.AddOrUpdate(key, new List<T> { value },
                (existingKey, existingValue) => { existingValue.Add(value); return existingValue; }
                );
        }

        public static void AddWinXToInterval(double winX, ConcurrentDictionary<int, long> intervals)
        {
            if (winX == 0)
            {
                intervals.AddOrUpdate(-1, 1, (key, oldValue) => oldValue + 1);
                return;
            }
            else if (winX > Intervals[Intervals.Count - 1])
            {
                int lastKey = Intervals[intervals.Count - 1];
                intervals.AddOrUpdate(lastKey, 1, (key, oldValue) => oldValue + 1);
                return;
            }
            for (int i = 1; i < Intervals.Count - 1; i++)
            {
                if (winX <= Intervals[i + 1])
                {
                    int key = Intervals[i];
                    intervals.AddOrUpdate(key, 1, (k, oldValue) => oldValue + 1);
                    break;
                }
            }
        }

        public static void ShowAvgWildsAmountInTheEndOfFSRound()
        {
            double avg = 0;
            foreach (var item in WildsAmountInTheEndOfFSRound)
            {
                avg += (double) item.Key / FSTriggersCount * item.Value ;
            }
            Console.WriteLine($"Avg Wilds in the end of FS = {avg}");
        }

        public static void ShowWinRate()
        {
            Console.WriteLine($"Win Rate >1x = {CurrentIteration / (double) WinRate}");
            Console.WriteLine($"Win Rate >10x = {CurrentIteration / (double) WinRateOf10x}");
            Console.WriteLine($"Win Rate >100x = {CurrentIteration / (double) WinRateOf100x}");
            Console.WriteLine($"Win Rate >1000x = {CurrentIteration / (double) WinRateOf1000x}");
        }

        public static void ShowHitRate()
        {
            Console.WriteLine($"Hit Rate = {CurrentIteration / (double) HitRate}");
        }

        public static void ShowROE()
        {
            double avg = 0;
            long sum = 0;
            foreach (var item in ROE)
            {
                avg += item.Key * item.Value;
                sum += item.Value;
            }
            avg /= sum;
            Console.WriteLine($"ROE = {avg}");
        }

        public static void CalculateStdDev()
        {
            double dispa = 0;
            foreach (var item in WinsPerTotalSpin)
            {
                double slagaemoe = (item.Key / CostToPlay - TotalRTP) * (item.Key / CostToPlay - TotalRTP) * item.Value / CurrentIteration;
                dispa += slagaemoe;
            }
            StdDev = Math.Pow(dispa, 0.5f);
        }

        public static void ShowAvgFSWinPerRound()
        {
            double avgWinPerFSRound = 0;
            foreach (var item in WinsPerFSRound)
            {
                avgWinPerFSRound += (double) item.Key / FSTriggersCount * item.Value;
            }
            avgWinPerFSRound /= CostToPlay;
            Console.WriteLine($"Avg win per FS round = {avgWinPerFSRound}x");
        }

        public static void CalculateTotalRTP()
        {
            TotalRTP = 0;
            foreach (var item in WinsPerTotalSpin)
            {
                TotalRTP += (double)item.Key / CurrentIteration * item.Value;
            }
            TotalRTP /= CostToPlay;
        }

        public static void CalculateBGOnlyRTP()
        {
            BGOnlyRTP = 0;
            foreach (var item in WinsPerBGOnlySpin)
            {
                BGOnlyRTP += (double)item.Key * item.Value;
            }
            BGOnlyRTP /= CurrentIteration * CostToPlay;
        }

        public static void CalculateFSOnlyRTP()
        {
            FreeSpinsRTP = 0;
            foreach (var item in WinsPerFSRound)
            {
                FreeSpinsRTP += (double)item.Key / CurrentIteration * item.Value / CostToPlay;
            }
        }

        public static void CalculateConfidenceInterval()
        {
            double temp = Math.Pow(CurrentIteration, 1.0 / 2);
            ConfidenceInterval = StdDev * 1.95 / temp; //95% CI
        }

        public static void ShowAvgFSCountPerFSTrigger()
        {
            Console.WriteLine($"Avg FS Count per trigger = {(double)FSCount / FSTriggersCount}");
        }

        public static void ShowRTPs()
        {
            Console.WriteLine($"Total RTP = {Math.Round(TotalRTP, 8) * 100}%");
            Console.WriteLine($"Total RTP CI 95% = {Math.Round((TotalRTP - ConfidenceInterval) * 100, 8)}% - {Math.Round((TotalRTP + ConfidenceInterval) * 100, 4)}%");
            Console.WriteLine($"BG only RTP = {Math.Round(BGOnlyRTP, 8) * 100}%");
            Console.WriteLine($"FS only RTP = {Math.Round(FreeSpinsRTP, 8) * 100}%");
        }

        public static void ShowStdDev()
        {
            double dispa = 0;
            foreach (var item in WinsPerTotalSpin)
            {
                double slagaemoe = (item.Key / CostToPlay - TotalRTP) * (item.Key / CostToPlay - TotalRTP) * item.Value;
                dispa += slagaemoe;
            }
            dispa /= CurrentIteration;
            Console.WriteLine($"Dispa = {Math.Round(dispa, 6)}");
            StdDev = Math.Sqrt(dispa);
            Console.WriteLine($"StdDev = {StdDev}");
        }

        public static void ShowIntervalsHitRate(ConcurrentDictionary<int, long> intervalToShow, string intervalName, ConcurrentDictionary<int, long> winsOfInterval)
        {
            long sumOfTheseSpins = (long) CurrentIteration;

            // Создаем отсортированный список пар ключ-значение из intervalToShow
            var sortedIntervals = intervalToShow.OrderBy(x => x.Key).ToList();

            Console.WriteLine($"{intervalName} Intervals Hit Rate (RTP%):");
            Console.WriteLine($"0x: {1 / ((double)sortedIntervals[0].Value / sumOfTheseSpins)} (0%)");

            double RTPofInterval = 0;
            for (int i = 1; i < sortedIntervals.Count - 1; i++)
            {
                RTPofInterval = GetRTPFromTheInterval(sortedIntervals[i].Key, sortedIntervals[i + 1].Key, winsOfInterval);
                Console.WriteLine($"{sortedIntervals[i].Key}x - {sortedIntervals[i + 1].Key}x: " +
                    $"{1 / ((double)sortedIntervals[i].Value / sumOfTheseSpins)} " +
                    $"({Math.Round(RTPofInterval * 100, 4)}%)");
            }
            RTPofInterval = GetRTPFromTheInterval(sortedIntervals.Last().Key, 999999999, winsOfInterval);
            Console.WriteLine($" > {sortedIntervals.Last().Key}x: " +
                $"{1 / ((double)sortedIntervals.Last().Value / sumOfTheseSpins)} ({Math.Round(RTPofInterval * 100, 4)}%)");
        }

        public static double GetRTPFromTheInterval(double leftBoardX, double rightBoardX, ConcurrentDictionary<int, long> winsOfInterval)
        {
            double result = 0;
            foreach (var item in winsOfInterval)
            {
                if (leftBoardX < item.Key / CostToPlay && item.Key / CostToPlay <= rightBoardX)
                {
                    result += (double) item.Key / CurrentIteration * item.Value ;
                }
            }
            result /= CostToPlay;
            return result;
        }

        public static void ShowScatterRTP()
        {
            Console.WriteLine($"RTP of scatters = {(double) FSTriggersCount / CurrentIteration * PayTable[Tomb][2] / 10 * 100}%");
        }

        public static void ShowSymbolsRTPs()
        {
            foreach (Symbol item in SymbolsHitsCount.Keys)
            {
                Console.Write($"{item}:     ");
                double RTPfromSymbol = 0;
                for (int i = 0; i <= 4; i++)
                {
                    RTPfromSymbol += SymbolsHitsCount[item][i] * PayTable[item][i] / CostToPlay / CurrentIteration * 100;
                }
                Console.WriteLine($"{Math.Round(RTPfromSymbol, 5)}%\t");
            }
        }

        public static void ShowSymbolsRTPs2()
        {
            Console.WriteLine($"{"Symbol",-12}{"2",-12}{"3",-12}{"4",-12}{"5",-12}");
            foreach (Symbol item in SymbolsHitsCount.Keys)
            {
                Console.Write($"{item + ":",-12}");
                double RTPfromSymbol = 0;
                for (int i = 2; i <= SlotWidth; i++)
                {
                    RTPfromSymbol = (double) SymbolsHitsCount[item][i - 1] / CurrentIteration * PayTable[item][i - 1] / CostToPlay * 100;
                    Console.Write($"{Math.Round(RTPfromSymbol, 2) + "%",-12}");
                }
                Console.WriteLine();
            }
        }

        public static void ShowFSTriggerCycle()
        {
            Console.WriteLine($"FS Cycle = {Math.Round((double)CurrentIteration / FSTriggersCount, 4)}");
        }

        public static void ShowMaxWinHitCycle()
        {
            if (WinsPerTotalSpin.ContainsKey(MAXWIN))
            {
                Console.WriteLine($"Max win Hit Rate = {(double) CurrentIteration / WinsPerTotalSpin[MAXWIN]}");
            }
            else
            {
                Console.WriteLine($"Max win Hit Rate = NaN");
            }
        }

        public static void ShowMaxWin()
        {
            MaxWin = 0;
            foreach (var item in WinsPerTotalSpin)
            {
                if (MaxWin < item.Key)
                {
                    MaxWin = item.Key;
                }
            }
            if (MaxWin > MAXWIN)
            {
                Console.WriteLine($"MaxWin = {MaxWin}");
                Console.ReadKey();
            }
            Console.WriteLine($"Max Win = {Math.Round(MaxWin / CostToPlay)}x");
        }

        public static void ShowMatrix(Symbol[,] matrix)
        {
            for (int i = 0; i < SlotHeight; i++)
            {
                for (int j = 0; j < SlotWidth; j++)
                {
                    string symbolToWrite = "";
                    symbolToWrite = matrix[i, j].ToString();
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

        public static void ShowStopPositions(List<int> stopPositions)
        {
            for (int j = 0; j < SlotWidth; j++)
            {
                string s = stopPositions[j].ToString();
                Console.Write($"{s}");
                for (int k = 0; k < 12 - s.Length; k++)
                {
                    Console.Write($" ");
                }
            }
            Console.WriteLine();
        }
    }

}
