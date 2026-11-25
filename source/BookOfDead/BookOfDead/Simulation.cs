using System;
using System.Collections.Generic;
using static DevilWest.Symbol;
using static DevilWest.SlotInfo;
using static DevilWest.SlotStats;
using System.Linq;
using OfficeOpenXml;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.Threading;
using System.Runtime.InteropServices;
using System.Collections.Concurrent;
using System.Diagnostics;

namespace DevilWest
{

    class Simulation
    {
        public static int ThreadsNumber = Environment.ProcessorCount;
        public static List<Thread> Threads =  new List<Thread>();
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ImportInfo();
            ResetStats();
            //Dictionary<string, Dictionary<int, List<Symbol>>> reelsetsAndTheirNames = new Dictionary<string, Dictionary<int, List<Symbol>>>();
            //reelsetsAndTheirNames.Add("Base Game Reels 1", BGReelsSets[0]);
            //reelsetsAndTheirNames.Add("Base Game Reels 2", BGReelsSets[1]);
            //reelsetsAndTheirNames.Add("Base Game Reels 3", BGReelsSets[2]);
            //foreach (var item in reelsetsAndTheirNames)
            //{
            //    ResetStats();
            //    for (int i = 1; i <= ThreadsNumber; i++)
            //    {
            //        ReelSetSimulator simulation = new ReelSetSimulator();
            //        Thread thread = new Thread(() => simulation.StartSimulationOfReelSet(i, item.Key, item.Value));
            //        thread.Priority = ThreadPriority.Highest;
            //        Threads.Add(thread);
            //        thread.Start();
            //        Thread.Sleep(100);
            //    }
            //    foreach (var item2 in Threads)
            //    {
            //        item2.Join();
            //    }
            //    Threads.Clear();
            //}
            for (int i = 1; i <= ThreadsNumber; i++)
            {
                SlotSimulation simulation = new SlotSimulation();
                Thread thread = new Thread(() => simulation.StartSimulationOfBG(i));
                thread.Priority = ThreadPriority.Highest;
                Threads.Add(thread);
                thread.Start();
                Thread.Sleep(30);
            }
            foreach (var item in Threads)
            {
                item.Join();
            }

            Console.ReadKey();
        }
    }

    public enum Symbol
    {
        Tomb,
        Men,
        Pharaon,
        Cat,
        Bird,
        A,
        K,
        Q,
        J,
        Ten,
        Blank,
        Inner,
    }

    class SlotSimulation
    {
        public Dictionary<int, List<Symbol>> CurrentBGReels = new Dictionary<int, List<Symbol>>();
        public Symbol[,] Matrix = new Symbol[SlotHeight, SlotWidth];
        public int[,] MultiplierMatrix = new int[SlotHeight, SlotWidth];
        public ThreadLocal<Random> Rand = new ThreadLocal<Random>(() => new Random(Guid.NewGuid().GetHashCode()));
        public List<int> StopPositions = new List<int>();
        public PaylinesCounter PaylinesCounter = new PaylinesCounter();
        public FreeSpinsFeature FreeSpinsFeature = new FreeSpinsFeature();
        private static Barrier _barrier1 = new Barrier(Simulation.ThreadsNumber);
        private static Barrier _barrier2 = new Barrier(Simulation.ThreadsNumber);
        private static int _temp = 0;
        private static Stopwatch _stopWatchInterval = new Stopwatch();
        public Symbol InnerSymbolAtCurrentSpin = default;

        public SlotSimulation()
        {
            FreeSpinsFeature.PaylinesCounter = PaylinesCounter;
            for (int i = 0; i < SlotWidth; i++)
            {
                StopPositions.Add(0);
            }
        }

        public void StartSimulationOfBG(int threadNumber) // lucky bet determined by int Config in SlotInfo.cs
        {
            using (var package = new ExcelPackage(new FileInfo($@"../../BookOfDeadIntervals.xlsx")))
            {
                using (var parsheetPackage = new ExcelPackage(new FileInfo($@"../../BookOfDeadParsheet.xlsx")))
                {
                    try
                    {
                        var workbook = package.Workbook;
                        var parsheetWorkbook = parsheetPackage.Workbook;
                        long intervalToUpdateStats = (int)10E6;
                        if (_stopWatchInterval.IsRunning == false)
                            _stopWatchInterval.Start();
                        for (; CurrentIteration <= IterationsCount + 1;)
                        {
                            if (CurrentIteration >= _temp * intervalToUpdateStats)
                            {
                                _barrier1.SignalAndWait();
                                if (threadNumber == Simulation.ThreadsNumber)
                                {
                                    _temp++;
                                    Console.Clear();
                                    Console.WriteLine("\x1b[3J");
                                    _stopWatchInterval.Stop();
                                    TimeSpan elapsedTime = _stopWatchInterval.Elapsed;
                                    Console.WriteLine($"Скорость симуляции {(int)(intervalToUpdateStats / elapsedTime.TotalSeconds)} BG spins/sec");
                                    _stopWatchInterval.Reset();
                                    _stopWatchInterval.Start();
                                    ShowStats();
                                    if (CurrentIteration % 10E6 == 0)
                                    {
                                        InsertIntervalsStats(workbook.Worksheets["Intervals"]);
                                        InsertSummaryStatsToParsheetOfBaseGame(parsheetWorkbook);
                                        package.Save();
                                        parsheetPackage.Save();
                                    }
                                }
                                _barrier2.SignalAndWait();
                            }
                            else
                            {
                                Interlocked.Increment(ref CurrentIteration);
                                MakeASpin(); //MakeASpinWithShowingDetails(); // запуск через кнопку  // запуск всего
                            }
                        }
                    }
                    finally
                    {
                        Console.WriteLine($"Поток {threadNumber} завершил работу");
                        Console.ReadKey();
                        if (threadNumber == Simulation.ThreadsNumber)
                        {
                            package.Save();
                            package.Dispose();
                        }
                    }
                }
            }
        }

        public void StartSimulationOfROE(int threadNumber)
        {
            try
            {
                long intervalToUpdateStats = (int)100;
                if (_stopWatchInterval.IsRunning == false)
                    _stopWatchInterval.Start();
                for (; CurrentIteration <= IterationsCount + 1;)
                {
                    if (CurrentIteration >= _temp * intervalToUpdateStats)
                    {
                        _barrier1.SignalAndWait();
                        if (threadNumber == Simulation.ThreadsNumber)
                        {
                            _temp++;
                            Console.Clear();
                            Console.WriteLine("\x1b[3J");
                            _stopWatchInterval.Stop();
                            TimeSpan elapsedTime = _stopWatchInterval.Elapsed;
                            Console.WriteLine($"Скорость симуляции {(int)(intervalToUpdateStats / elapsedTime.TotalSeconds)} BG spins/sec");
                            _stopWatchInterval.Reset();
                            _stopWatchInterval.Start();
                            ShowROE();
                            Console.ReadKey();
                        }
                        _barrier2.SignalAndWait();
                    }
                    else
                    {
                        Interlocked.Increment(ref CurrentIteration);
                        MakeROESpin();
                    }
                }
            }
            finally
            {
                Console.WriteLine($"Поток {threadNumber} завершил работу");
                Console.ReadKey();
            }

        }

        public void StartSimulationOfFS(int threadNumber)
        {
            if (threadNumber == 1)
                CostToPlay *= 100;
            using (var package = new ExcelPackage(new FileInfo($@"../../BookOfDeadIntervals.xlsx")))
            {
                using (var parsheetPackage = new ExcelPackage(new FileInfo($@"../../BookOfDeadParsheet.xlsx")))
                {
                    try
                    {
                        var workbook = package.Workbook;
                        var parsheetWorkbook = parsheetPackage.Workbook;
                        long intervalToUpdateStats = (int)1E6;
                        if (_stopWatchInterval.IsRunning == false)
                            _stopWatchInterval.Start();
                        for (; CurrentIteration <= IterationsCount + 1;)
                        {
                            if (CurrentIteration >= _temp * intervalToUpdateStats)
                            {
                                _barrier1.SignalAndWait();
                                if (threadNumber == Simulation.ThreadsNumber)
                                {
                                    _temp++;
                                    Console.Clear();
                                    Console.WriteLine("\x1b[3J");
                                    _stopWatchInterval.Stop();
                                    TimeSpan elapsedTime = _stopWatchInterval.Elapsed;
                                    Console.WriteLine($"Скорость симуляции {(int)(intervalToUpdateStats / elapsedTime.TotalSeconds)} BG spins/sec");
                                    _stopWatchInterval.Reset();
                                    _stopWatchInterval.Start();
                                    ShowStats();
                                    if (CurrentIteration % 1E6 == 0)
                                    {
                                        InsertSummaryStatsToParsheetOfBuyFeature(parsheetWorkbook);
                                        InsertIntervalsStats(workbook.Worksheets["Intervals"]);
                                        package.Save();
                                        parsheetPackage.Save();
                                    }
                                }
                                _barrier2.SignalAndWait();
                            }
                            else
                            {
                                Interlocked.Increment(ref CurrentIteration);
                                MakeFSRound();
                            }
                        }
                    }
                    finally
                    {
                        Console.WriteLine($"Поток {threadNumber} завершил работу");
                        Console.ReadKey();
                        if (threadNumber == Simulation.ThreadsNumber)
                        {
                            package.Save();
                            package.Dispose();
                        }
                    }
                }
            }
        }

        private void InsertMAXWINHitRate(ExcelWorksheet worksheet, string cell)
        {          
            if (WinsPerTotalSpin.TryGetValue(MAXWIN, out long count) == false)
            {
                worksheet.Cells[cell].Value = "NaN";
            }
            else
            {
                worksheet.Cells[cell].Value = (double)CurrentIteration / count;
            }
        }

        private void InsertSummaryStatsToParsheetOfBaseGame(ExcelWorkbook workbook)
        {
            var summaryWorksheet = workbook.Worksheets["Summary"];
            char column = 'C';
            if (Config == 2)
                column = 'G';
            //InsertFSHitRate(summaryWorksheet, 12, (char) (column + 1));
            InsertRTPofBG(summaryWorksheet, $"{column}11");
            InsertRTPofFS(summaryWorksheet, $"{column}12");
            InsertStdDev(summaryWorksheet, 16, column);
            InsertWinRate(summaryWorksheet, $"{column}17");
            Insert10xRate(summaryWorksheet, $"{column}18");
            Insert100xRate(summaryWorksheet, $"{column}19");
            Insert1000xRate(summaryWorksheet, $"{column}20");
            InsertHitRate(summaryWorksheet, $"{column}21");
            InsertMaxWinHitRate(summaryWorksheet, $"{column}25");
        }

        private void InsertSummaryStatsToParsheetOfBuyFeature(ExcelWorkbook workbook)
        {
            var summaryWorksheet = workbook.Worksheets["Summary"];
            InsertRTPofBG(summaryWorksheet, "K11");
            InsertRTPofFS(summaryWorksheet, "K12");
            InsertStdDev(summaryWorksheet, 16, 'K');
            InsertWinRate(summaryWorksheet, "K17");
            Insert10xRate(summaryWorksheet, "K18");
            Insert100xRate(summaryWorksheet, "K19");
            Insert1000xRate(summaryWorksheet, "K20");
            InsertHitRate(summaryWorksheet, "K21");
            InsertMaxWinHitRate(summaryWorksheet, "K25");
        }


        private void InsertIntervalsStats(ExcelWorksheet worksheet)
        {
            InsertIntervalToSheetOfStats(worksheet, IntervalBaseGameWinsX, WinsPerBGOnlySpin, "C", "D");
            InsertIntervalToSheetOfStats(worksheet, IntervalFSRoundWinsX, WinsPerFSRound, "M", "N");
            InsertIntervalToSheetOfStats(worksheet, IntervalTotalWinsX, WinsPerTotalSpin, "W", "X");
            InsertAvgFSCount(worksheet, 33, 'M');
            InsertStdDev(worksheet, 34, 'M');
            InsertFSHitRate(worksheet, 35, 'M');
            InsertCurrentIteration(worksheet, 1, 'A');
        }

        private void InsertMaxWinHitRate(ExcelWorksheet worksheet, string cell)
        {
            if (WinsPerTotalSpin.ContainsKey(MAXWIN))
            {
                worksheet.Cells[cell].Value = ((double)CurrentIteration / WinsPerTotalSpin[MaxWin]);
            }
            else
            {
                worksheet.Cells[cell].Value = "NaN";
            }
        }

        private void InsertHitRate(ExcelWorksheet worksheet, string cell)
        {
            worksheet.Cells[cell].Value = (double)CurrentIteration / HitRate;
        }

        private void InsertWinRate(ExcelWorksheet worksheet, string cell)
        {
            worksheet.Cells[cell].Value = (double)CurrentIteration / WinRate;
        }

        private void Insert10xRate(ExcelWorksheet worksheet, string cell)
        {
            worksheet.Cells[cell].Value = (double)CurrentIteration / WinRateOf10x;
        }

        private void Insert100xRate(ExcelWorksheet worksheet, string cell)
        {
            worksheet.Cells[cell].Value = (double)CurrentIteration / WinRateOf100x;
        }

        private void Insert1000xRate(ExcelWorksheet worksheet, string cell)
        {
            worksheet.Cells[cell].Value = (double)CurrentIteration / WinRateOf1000x;
        }

        private void InsertRTPofBG(ExcelWorksheet worksheet, string cell)
        {
            worksheet.Cells[cell].Value = BGOnlyRTP;
        }

        private void InsertRTPofFS(ExcelWorksheet worksheet, string cell)
        {
            worksheet.Cells[cell].Value = FreeSpinsRTP;
        }

        private void InsertCurrentIteration(ExcelWorksheet worksheet, int row, char column)
        {
            worksheet.Cells[$"{column}{row}"].Value = $"Iteration: {CurrentIteration}";
        }

        private void InsertFSHitRate(ExcelWorksheet worksheet, int row, char column)
        {
            worksheet.Cells[$"{column}{row}"].Value = (double)CurrentIteration / FSTriggersCount;
        }

        private void InsertStdDev(ExcelWorksheet worksheet, int row, char column)
        {
            worksheet.Cells[$"{column}{row}"].Value = StdDev;
        }

        private void InsertAvgFSCount(ExcelWorksheet worksheet, int row, char column)
        {
            worksheet.Cells[$"{column}{row}"].Value = (double) FSCount / FSTriggersCount;
        }

        public void InsertIntervalToSheetOfStats(ExcelWorksheet worksheet, ConcurrentDictionary<int, long> interval, ConcurrentDictionary<int, long> winsOfInterval, string columnForRTP, string columnForHitRate, int startRow = 4)
        {
            long sumOfTheseSpins = CurrentIteration;

            // Создаем отсортированный список пар ключ-значение из interval
            var sortedIntervals = interval.OrderBy(x => x.Key).ToList();

            worksheet.Cells[$"{columnForHitRate}{startRow}"].Value = 1 / ((double)sortedIntervals[0].Value / sumOfTheseSpins);
            worksheet.Cells[$"{columnForRTP}{startRow}"].Value = 0;
            double RTPofInterval = 0;

            for (int i = 1; i < sortedIntervals.Count - 1; i++)
            {
                RTPofInterval = GetRTPFromTheInterval(sortedIntervals[i].Key, sortedIntervals[i + 1].Key, winsOfInterval);
                worksheet.Cells[$"{columnForRTP}{startRow + i}"].Value = RTPofInterval;
                worksheet.Cells[$"{columnForHitRate}{startRow + i}"].Value = 1 / ((double)sortedIntervals[i].Value / sumOfTheseSpins);
            }

            RTPofInterval = GetRTPFromTheInterval(sortedIntervals.Last().Key, 999999999, winsOfInterval);
            worksheet.Cells[$"{columnForRTP}{startRow + sortedIntervals.Count - 1}"].Value = RTPofInterval;
            worksheet.Cells[$"{columnForHitRate}{startRow + sortedIntervals.Count - 1}"].Value = 1 / ((double)sortedIntervals.Last().Value / sumOfTheseSpins);
        }

        private void ShowStats()
        {
            Console.WriteLine($"Iteration: {CurrentIteration}");
            CalculateTotalRTP();
            CalculateBGOnlyRTP();
            CalculateFSOnlyRTP();
            CalculateStdDev();
            CalculateConfidenceInterval();
            ShowFSTriggerCycle();
            ShowRTPs();
            Console.WriteLine();
            ShowAvgFSWinPerRound();
            ShowAvgFSCountPerFSTrigger();
            ShowAvgWildsAmountInTheEndOfFSRound();
            Console.WriteLine();
            Console.WriteLine();
            ShowStdDev();
            ShowMaxWin();
            ShowMaxWinHitCycle();
            ShowWinRate();
            ShowHitRate();
            //Console.WriteLine();
            //ShowBGRTPsFromReelSets();
            //Console.WriteLine();

            Console.WriteLine();
            ShowSymbolsRTPs2();
            ShowScatterRTP();
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine();

            //ShowIntervalsHitRate(IntervalBaseGameWinsX, "Base Game Wins", WinsPerBGOnlySpin);
            //Console.WriteLine();
            //ShowIntervalsHitRate(IntervalFSRoundWinsX, "Free Spins Round Wins", WinsPerFSRound);
            //Console.WriteLine();
            //ShowIntervalsHitRate(IntervalTotalWinsX, "Total Wins", WinsPerTotalSpin);
            //Console.WriteLine();
        }

        public void GenerateNewMatrix(Dictionary<int, List<Symbol>> reelSet)
        {
            GenerateNewStopPositions(reelSet);
            for (int i = 0; i < SlotWidth; i++)
            {
                for (int j = 0; j < SlotHeight; j++)
                {
                    Matrix[j, i] = reelSet[i][(StopPositions[i] + j) % reelSet[i].Count];
                }
            }
        }

        private void MakeFSRound()
        {
            int totalWin = 0;
            int symbolsWins = 0;
            int scattersAmount = 0;
            int FSRoundWin = 0;
            GenerateNewMatrix(ReelSetOfBuyFeature);
            scattersAmount = GetSymbolCountFromMatrix(Tomb);
            symbolsWins = PaylinesCounter.GetPaylinesWins(Matrix, MultiplierMatrix);
            if (symbolsWins < MAXWIN)
            {
                if (scattersAmount >= 3)
                {
                    symbolsWins += PayTable[Tomb][scattersAmount - 1];
                    Interlocked.Increment(ref FSTriggersCount);

                    FreeSpinsFeature.StartFeature(out FSRoundWin, scattersAmount);
                    if (FSRoundWin + symbolsWins >= MAXWIN)
                    {
                        symbolsWins = MAXWIN - FSRoundWin;
                    }
                    AddWinTo(FSRoundWin, WinsPerFSRound);
                    AddWinXToInterval(FSRoundWin / CostToPlay, IntervalFSRoundWinsX);
                }
            }
            else
            {
                symbolsWins = MAXWIN;
            }

            totalWin = symbolsWins + FSRoundWin;
            //Console.WriteLine($"Total Win = {totalWin}");
            //Console.ReadKey();
            if (totalWin > 0)
            {
                Interlocked.Increment(ref HitRate);
                if (totalWin > CostToPlay)
                    Interlocked.Increment(ref WinRate);
                if (totalWin > CostToPlay / 10)
                    Interlocked.Increment(ref WinRateOf10x);
                if (totalWin > CostToPlay)
                    Interlocked.Increment(ref WinRateOf100x);
                if (totalWin > CostToPlay * 10)
                    Interlocked.Increment(ref WinRateOf1000x);
            }
            AddWinTo(totalWin, WinsPerTotalSpin);
            AddWinTo(symbolsWins, WinsPerBGOnlySpin);
            AddWinXToInterval(totalWin / CostToPlay, IntervalTotalWinsX);
            AddWinXToInterval(symbolsWins / CostToPlay, IntervalBaseGameWinsX);
        }

        private void MakeASpin()
        {
            int totalWin = 0;
            int symbolsWins = 0;
            int scattersAmount = 0;
            int FSRoundWin = 0;
            // CurrentBGReels = TableWeightsSelector.GetRandomObjectFromTableWithWeights(BGReelsWeights, Rand);
            CurrentBGReels = BGReelsSets[0];
            GenerateNewMatrix(CurrentBGReels);
            RealizeMultiplierMatrix();
            SetInnersSymbolsDuringBG();
            RealizeInnerSymbols();
            scattersAmount = GetSymbolCountFromMatrix(Tomb);
            symbolsWins = PaylinesCounter.GetPaylinesWins(Matrix, MultiplierMatrix);
            if (symbolsWins < MAXWIN)
            {
                if (scattersAmount >= 3)   // need be 3
                {
                    symbolsWins += PayTable[Tomb][scattersAmount - 1];
                    Interlocked.Increment(ref FSTriggersCount);

                    FreeSpinsFeature.StartFeature(out FSRoundWin, scattersAmount);
                    if (FSRoundWin + symbolsWins >= MAXWIN)
                    {
                        symbolsWins = MAXWIN - FSRoundWin;
                    }
                    AddWinTo(FSRoundWin, WinsPerFSRound);
                    AddWinXToInterval(FSRoundWin / CostToPlay, IntervalFSRoundWinsX);
                }
            }
            else
            {
                symbolsWins = MAXWIN;
            }

            totalWin = symbolsWins + FSRoundWin;
            AddWinTo(symbolsWins, WinsFromReelSetsBGOnly[CurrentBGReels]);
            //Console.WriteLine($"Total Win = {totalWin}");
            //Console.ReadKey();
            if (totalWin > 0)
            {
                Interlocked.Increment(ref HitRate);
                if (totalWin > CostToPlay)
                    Interlocked.Increment(ref WinRate);
                if (totalWin > CostToPlay * 10)
                    Interlocked.Increment(ref WinRateOf10x);
                if (totalWin > CostToPlay * 100)
                    Interlocked.Increment(ref WinRateOf100x);
                if (totalWin > CostToPlay * 1000)
                    Interlocked.Increment(ref WinRateOf1000x);
            }
                
            AddWinTo(totalWin, WinsPerTotalSpin);
            AddWinTo(symbolsWins, WinsPerBGOnlySpin);
            AddWinXToInterval(totalWin / CostToPlay, IntervalTotalWinsX);
            AddWinXToInterval(symbolsWins / CostToPlay, IntervalBaseGameWinsX);
        }


        private void MakeASpinWithShowingDetails()
        {
            int totalWin = 0;
            int symbolsWins = 0;
            int scattersAmount = 0;
            int FSRoundWin = 0;
            CurrentBGReels = BGReelsSets[0];
            GenerateNewMatrix(CurrentBGReels);
            RealizeMultiplierMatrix();
            SetInnersSymbolsDuringBG();
            RealizeInnerSymbols();
            ShowMatrix(Matrix);
            symbolsWins = PaylinesCounter.GetPaylinesWinsWithShowingDetails(Matrix, MultiplierMatrix);
            scattersAmount = GetSymbolCountFromMatrix(Tomb);
            if (symbolsWins < MAXWIN)
            {
                if (scattersAmount == 0)
                {
                    Interlocked.Increment(ref FSTriggersCount);
                    Console.WriteLine("Free Spins triggered!");
                    FreeSpinsFeature.StartFeatureWithShowingDetails(out FSRoundWin, 4);
                    if (FSRoundWin + symbolsWins >= MAXWIN)
                    {
                        symbolsWins = MAXWIN - FSRoundWin;
                    }
                    AddWinTo(FSRoundWin, WinsPerFSRound);
                    AddWinXToInterval(FSRoundWin / CostToPlay, IntervalFSRoundWinsX);
                }
            }
            else
            {
                symbolsWins = MAXWIN;
            }
            Console.ReadKey();
            totalWin = symbolsWins + FSRoundWin;
            AddWinTo(symbolsWins, WinsFromReelSetsBGOnly[CurrentBGReels]);
            //Console.WriteLine($"Total Win = {totalWin}");
            //Console.ReadKey();
            if (totalWin > CostToPlay)
                Interlocked.Increment(ref WinRate);
            if (totalWin > 0)
                Interlocked.Increment(ref HitRate);
            AddWinTo(totalWin, WinsPerTotalSpin);
            AddWinTo(symbolsWins, WinsPerBGOnlySpin);
            AddWinXToInterval(totalWin / CostToPlay, IntervalTotalWinsX);
            AddWinXToInterval(symbolsWins / CostToPlay, IntervalBaseGameWinsX);
        }

        public void MakeROESpin()
        {
            int balance = 100 * (int) CostToPlay;
            int spinsCount = 0;
            while (balance >= CostToPlay)
            {
                spinsCount++;
                balance -= (int)CostToPlay;
                int totalWin = 0;
                int symbolsWins = 0;
                int scattersAmount = 0;
                int FSRoundWin = 0; ;
                CurrentBGReels = BGReelsSets[0];
                GenerateNewMatrix(CurrentBGReels);
                symbolsWins = PaylinesCounter.GetPaylinesWins(Matrix, MultiplierMatrix);

                scattersAmount = GetSymbolCountFromMatrix(Tomb);
                if (scattersAmount == 3)
                {
                    Interlocked.Increment(ref FSTriggersCount);

                    FreeSpinsFeature.StartFeature(out FSRoundWin, scattersAmount);
                    if (FSRoundWin / (float)Paylines.Count > MAXWIN)
                    {
                        FSRoundWin = MAXWIN * Paylines.Count - symbolsWins;
                    }
                    AddWinTo(FSRoundWin, WinsPerFSRound);
                    AddWinXToInterval(FSRoundWin / CostToPlay, IntervalFSRoundWinsX);
                }


                totalWin = symbolsWins + FSRoundWin;
                if (totalWin >= CostToPlay)
                    Interlocked.Increment(ref WinRate);
                if (totalWin > 0)
                    Interlocked.Increment(ref HitRate);
                AddWinTo(totalWin, WinsPerTotalSpin);
                AddWinTo(symbolsWins, WinsPerBGOnlySpin);
                AddWinXToInterval(totalWin / CostToPlay, IntervalTotalWinsX);
                AddWinXToInterval(symbolsWins / CostToPlay, IntervalBaseGameWinsX);
                balance += totalWin;
            }
            AddWinTo(spinsCount, ROE);
        }

        public void GenerateNewStopPositions(Dictionary<int, List<Symbol>> reelSet)
        {
            for (int i = 0; i < SlotWidth; i++)
            {
                StopPositions[i] = Rand.Value.Next(0, reelSet[i].Count);
            }
        }

        public int GetSymbolCountFromMatrix(Symbol symbol)
        {
            int result = 0;
            for (int i = 0; i < SlotWidth; i++)
            {
                for (int j = 0; j < SlotHeight; j++)
                {
                    if (Matrix[j, i] == symbol)
                    {
                        result++;
                    }
                    
                }
            }
            return result;
        }

        public void RealizeMultiplierMatrix()
        {
            for (int i = 0; i < SlotHeight; i++)
            {
                for (int j = 1; j < SlotWidth - 1; j++)
                {
                    if (Matrix[i, j] == Tomb)
                    {
                        MultiplierMatrix[i, j] = TableWeightsSelector.GetRandomObjectFromTableWithWeights(WildsMultipliersOfBG, Rand);
                    }
                    else
                    {
                        MultiplierMatrix[i, j] = 1;
                    }
                }
            }
        }

        public void SetInnersSymbolsDuringBG()
        {
            InnerSymbolAtCurrentSpin = TableWeightsSelector.GetRandomObjectFromTableWithWeights(InnerReelSetsBG[CurrentBGReels], Rand);
        }

        public void RealizeInnerSymbols()
        {
            for (int i = 0; i < SlotWidth; i++)
            {
                for (int j = 0; j < SlotHeight; j++)
                {
                    if (Matrix[j, i] == Inner)
                    {
                        Matrix[j, i] = InnerSymbolAtCurrentSpin;
                    }
                }
            }
        }
    }
}