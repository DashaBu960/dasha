using System;  // Базовые типы .NET (Random, Guid и т.д.)
using System.Collections.Generic;  // Коллекции: Dictionary, List
using static DevilWest.Symbol;  // Статический импорт символов (Wild, Blank, Inner...) без префикса Symbol.
using static DevilWest.SlotInfo;  // Статический импорт настроек слота (SlotWidth, SlotHeight и т.д.)
using static DevilWest.SlotStats;  // Статический импорт статистики (счётчики, AddWinTo и т.д.)
using System.Linq;  // LINQ - удобные запросы к коллекциям
using OfficeOpenXml;  // EPPlus - работа с Excel-файлами (ExcelPackage)
using System.IO;  // Работа с файлами (FileInfo)
using System.Reflection;  // Рефлексия (доступ к метаданным типов)
using System.Windows.Media.Imaging;  // Типы изображений WPF (здесь, по-видимому, не используется)
using System.Threading;  // Многопоточность: Thread, ThreadLocal, Barrier, Interlocked
using System.Runtime.InteropServices;  // Interop (здесь, по-видимому, не используется)
using System.Collections.Concurrent;  // Потокобезопасные коллекции
using System.Diagnostics;  // Stopwatch - замер времени

namespace DevilWest  // Пространство имён проекта
{
    class ReelSetSimulator  // Класс симулятора одного набора барабанов (виден только внутри сборки)
    {
        public Dictionary<int, List<Symbol>> CurrentBGReels = new Dictionary<int, List<Symbol>>();  // Текущие барабаны основной игры: номер барабана -> список символов
        public Symbol[,] Matrix = new Symbol[SlotHeight, SlotWidth];  // Игровое поле: символы по строкам и столбцам
        public int[,] MultiplierMatrix = new int[SlotHeight, SlotWidth];  // Множители для каждой ячейки поля
        public ThreadLocal<Random> Rand = new ThreadLocal<Random>(() => new Random(Guid.NewGuid().GetHashCode()));  // Свой Random на каждый поток; seed из Guid, чтобы потоки не повторяли числа
        public List<int> StopPositions = new List<int>();  // Позиции остановки каждого барабана
        public PaylinesCounter PaylinesCounter = new PaylinesCounter();  // Подсчёт выигрышей по линиям выплат
        public FreeSpinsFeature FreeSpinsFeature = new FreeSpinsFeature();  // Бонусная игра с фриспинами
        private static Barrier _barrier1 = new Barrier(Simulation.ThreadsNumber);  // Барьер 1: потоки ждут друг друга в этой точке (число потоков = ThreadsNumber), общий для всех экземпляров
        private static Barrier _barrier2 = new Barrier(Simulation.ThreadsNumber);  // Барьер 2: вторая точка синхронизации потоков
        private static int _temp = 0;  // Общий счётчик-временная переменная для всех потоков
        private static Stopwatch _stopWatchInterval = new Stopwatch();  // Общий секундомер для замера интервалов между выводами статистики
        public Symbol InnerSymbolAtCurrentSpin = default;  // Символ, в который превращается Inner на текущем спине
        private static string _namesOfDoneReelSets = "";  // Имена уже просчитанных наборов барабанов (общая строка для всех потоков)

        public ReelSetSimulator()
        {
            FreeSpinsFeature.PaylinesCounter = PaylinesCounter;
            for (int i = 0; i < SlotWidth; i++)
            {
                StopPositions.Add(0);
            }
        }

        public void StartSimulationOfReelSet(int threadNumber, string reelSetName, Dictionary<int, List<Symbol>> reelSet) // lucky bet determined by int Config in SlotInfo.cs
        {
            using (var parsheetPackage = new ExcelPackage(new FileInfo($@"../../DevilWestParsheet.xlsx")))
            {
                try
                {
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
                                Console.WriteLine($"Working with {reelSetName}");
                                _stopWatchInterval.Stop();
                                TimeSpan elapsedTime = _stopWatchInterval.Elapsed;
                                Console.WriteLine($"Скорость симуляции {(int)(intervalToUpdateStats / elapsedTime.TotalSeconds)} BG spins/sec");
                                _stopWatchInterval.Reset();
                                _stopWatchInterval.Start();
                                ShowStats();
                                Console.WriteLine($"Done reelSets:\n{_namesOfDoneReelSets}");
                                if (CurrentIteration % 10E6 == 0)
                                {
                                    InsertSummaryStatsToParsheetOfReelSet(parsheetWorkbook, reelSetName);
                                    parsheetPackage.Save();
                                }
                            }
                            _barrier2.SignalAndWait();
                        }
                        else
                        {
                            Interlocked.Increment(ref CurrentIteration);
                            MakeASpin(reelSet);
                        }
                    }
                }
                finally
                {
                    Console.WriteLine($"Поток {threadNumber} завершил работу");
                    if (threadNumber == Simulation.ThreadsNumber)
                    {
                        InsertSummaryStatsToParsheetOfReelSet(parsheetPackage.Workbook, reelSetName);
                        parsheetPackage.Save();
                        parsheetPackage.Dispose(); 
                        _namesOfDoneReelSets += reelSetName + '\n';
                        _temp = 0;
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

        private void InsertSummaryStatsToParsheetOfReelSet(ExcelWorkbook workbook, string reelSetName)
        {
            var summaryWorksheet = workbook.Worksheets[reelSetName];
            char column = 'L';
            if (Config == 2)
                column = 'M';
            InsertRTPofBG(summaryWorksheet, $"{column}2");
            InsertHitRate(summaryWorksheet, $"{column}3");
            InsertWinRate(summaryWorksheet, $"{column}4");
            InsertWinRateForLuckyBet(summaryWorksheet, $"M4");
        }

        private void InsertHitRate(ExcelWorksheet worksheet, string cell)
        {
            worksheet.Cells[cell].Value = (double)CurrentIteration / HitRate;
        }

        private void InsertWinRate(ExcelWorksheet worksheet, string cell)
        {
            worksheet.Cells[cell].Value = (double)CurrentIteration / WinRate;
        }

        private void InsertWinRateForLuckyBet(ExcelWorksheet worksheet, string cell)
        {
            worksheet.Cells[cell].Value = (double)CurrentIteration / WinRateOfLuckyBet;
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
            worksheet.Cells[$"{column}{row}"].Value = (double)FSCount / FSTriggersCount;
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
            Console.WriteLine();
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

        private void MakeASpin(Dictionary<int, List<Symbol>> reelSet)
        {
            int totalWin = 0;
            int symbolsWins = 0;
            int scattersAmount = 0;
            int FSRoundWin = 0;
            CurrentBGReels = reelSet;
            GenerateNewMatrix(CurrentBGReels);
            //RealizeMultiplierMatrix();
            SetInnersSymbolsDuringBG();
            RealizeInnerSymbols();
            scattersAmount = GetSymbolCountFromMatrix(Scatter);
            symbolsWins = PaylinesCounter.GetPaylinesWins(Matrix, MultiplierMatrix);
            if (symbolsWins < MAXWIN)
            {
                if (scattersAmount >= 3)
                {
                    symbolsWins += PayTable[Scatter][scattersAmount - 1];
                    Interlocked.Increment(ref FSTriggersCount);
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
                if (totalWin > CostToPlay * 1.5f)
                    Interlocked.Increment(ref WinRateOfLuckyBet);
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
            scattersAmount = GetSymbolCountFromMatrix(Scatter);
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
            int balance = 100 * (int)CostToPlay;
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

                scattersAmount = GetSymbolCountFromMatrix(Wild);
                if (scattersAmount == 3)
                {
                    Interlocked.Increment(ref FSTriggersCount);

                    FreeSpinsFeature.StartFeature(out FSRoundWin, scattersAmount);
                    if (FSRoundWin / 20.0f > MAXWIN)
                    {
                        FSRoundWin = MAXWIN * 20 - symbolsWins;
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
                    if (Matrix[i, j] == Wild)
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
