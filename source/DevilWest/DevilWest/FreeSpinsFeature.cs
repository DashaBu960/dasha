using System;
using System.Collections.Generic;
using static DevilWest.Symbol;
using static DevilWest.SlotInfo;
using static DevilWest.SlotStats;
using static DevilWest.TableWeightsSelector;
using System.Linq;
using OfficeOpenXml;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.Threading;

namespace DevilWest
{
    public class FreeSpinsFeature
    {
        public Dictionary<int, List<Symbol>> CurrentFSReels = FSReelsSets[0];
        public PaylinesCounter PaylinesCounter = new PaylinesCounter();
        public ThreadLocal<Random> Rand = new ThreadLocal<Random>(() => new Random(Guid.NewGuid().GetHashCode()));
        public List<int> StopPositions = new List<int> { 0, 0, 0, 0, 0, 0 };
        public Symbol InnerSymbolAtCurrentSpin = default;

        private Symbol[,] _matrix = new Symbol[SlotHeight, SlotWidth];
        private int[,] _multiplierMatrix = new int[SlotHeight, SlotWidth];
        private Symbol[,] _wildMatrix = new Symbol[SlotHeight, SlotWidth];
        private List<(int, int)> _positionsOfMultiplierMatrixNeedsToBeClearedAfterSpin = new List<(int, int)>();

        private Symbol _activatedSymbol = default;
        private int _currentMultiplierOfChosenSymbol = 1;
        private int _freeSpinsAmount = 0;
        private int _totalWin;
        private int _currentFSNumber = 0;


        public void StartFeature(out int totalWin, int scattersAmount)
        {
            _positionsOfMultiplierMatrixNeedsToBeClearedAfterSpin.Clear();
            ClearMatrix();
            ClearMultiplierMatrix();
            _totalWin = 0;
            _activatedSymbol = TableWeightsSelector.GetRandomObjectFromTableWithWeights(WeightsToChooseSymbol, Rand);
            _currentMultiplierOfChosenSymbol = TableWeightsSelector.GetRandomObjectFromTableWithWeights(ChosenSymbolsMultipliers, Rand);
            _freeSpinsAmount = StartFreeSpinsAmount[scattersAmount];
            _freeSpinsAmount += TableWeightsSelector.GetRandomObjectFromTableWithWeights(AdditionalFSAmountsTable, Rand);
            _currentFSNumber = 0;
            for (_currentFSNumber = 0; _currentFSNumber < _freeSpinsAmount; _currentFSNumber++)
            {
                MakeASpin(out int win);
                if (win + _totalWin >= MAXWIN)
                {
                    win = MAXWIN - _totalWin;
                    _totalWin = MAXWIN;
                    AddWinTo(win, WinsPerFreeSpin);
                    break;
                }
                else
                {
                    AddWinTo(win, WinsPerFreeSpin);
                    _totalWin += win;
                }
            }
            Interlocked.Add(ref FSCount, _freeSpinsAmount);
            totalWin = _totalWin;
            if (totalWin >= MAXWIN)
            {
                totalWin = MAXWIN;
            }
            AddWinTo(GetSymbolCountFromMatrix(Wild), WildsAmountInTheEndOfFSRound);
        }


        public void StartFeatureWithShowingDetails(out int totalWin, int scattersAmount)
        {
            Console.Clear();
            _positionsOfMultiplierMatrixNeedsToBeClearedAfterSpin.Clear();
            ClearMatrix();
            ClearMultiplierMatrix();
            _totalWin = 0;
            _activatedSymbol = TableWeightsSelector.GetRandomObjectFromTableWithWeights(WeightsToChooseSymbol, Rand);
            _currentMultiplierOfChosenSymbol = TableWeightsSelector.GetRandomObjectFromTableWithWeights(ChosenSymbolsMultipliers, Rand);
            Console.WriteLine($"Current mult of {_activatedSymbol} = {_currentMultiplierOfChosenSymbol}");
            int startFreeSpinsAmount = StartFreeSpinsAmount[scattersAmount];
            int additionalFreeSpinsAmount = TableWeightsSelector.GetRandomObjectFromTableWithWeights(AdditionalFSAmountsTable, Rand);
            _freeSpinsAmount = startFreeSpinsAmount + additionalFreeSpinsAmount;
            _currentFSNumber = 0;
            Console.WriteLine($"{startFreeSpinsAmount} + {additionalFreeSpinsAmount} = {_freeSpinsAmount}FS are awarded");
            for (_currentFSNumber = 0; _currentFSNumber < _freeSpinsAmount; _currentFSNumber++)
            {
                Console.WriteLine($"Spin #{_currentFSNumber}");
                MakeASpinWithShowingDetails(out int win);
                if (win + _totalWin >= MAXWIN)
                {
                    win = MAXWIN - _totalWin;
                    _freeSpinsAmount = 0;
                    _totalWin = MAXWIN;
                    AddWinTo(win, WinsPerFreeSpin);
                    break;
                }
                else
                {
                    AddWinTo(win, WinsPerFreeSpin);
                    _totalWin += win;
                }
            }
            FSCount += _freeSpinsAmount;
            totalWin = _totalWin;
            if (totalWin >= MAXWIN)
            {
                totalWin = MAXWIN;
            }
        }

        private void MakeASpin(out int win)
        {
            win = 0;
            ChooseCurrentFSReelSet();
            GenerateNewMatrix(CurrentFSReels);
            SetInnersSymbolsDuringFS();
            RealizeInnerSymbols();
          //  RealizeMultiplierMatrix();
            int paylinesWin = PaylinesCounter.GetPaylinesWins(_matrix, _multiplierMatrix);
            // выбираем множитель этого фриспина по весам из таблицы FSMultipliersTable (лист "FS Multipliers")
            int fsMultiplier = TableWeightsSelector.GetRandomObjectFromTableWithWeights(FSMultipliersTable, Rand);
            // итоговый выигрыш спина = выигрыш по линиям * множитель фриспина
            win = paylinesWin * fsMultiplier;
            ClearMultiplierMatrixOutOfSymbolsMulties();
        }

        private void MakeASpinWithShowingDetails(out int win)
        {
            win = 0;
            ChooseCurrentFSReelSet();
            GenerateNewMatrix(CurrentFSReels);
            SetInnersSymbolsDuringFS();
            RealizeInnerSymbols();
            RealizeMultiplierMatrix();
            ShowMatrix(_matrix);
            int paylinesWin = PaylinesCounter.GetPaylinesWinsWithShowingDetails(_matrix, _multiplierMatrix);
            if (paylinesWin > 0)
            {
                ShowMultiplierMatrix();
            }
            // выбираем множитель этого фриспина по весам из таблицы FSMultipliersTable (лист "FS Multipliers")
            int fsMultiplier = TableWeightsSelector.GetRandomObjectFromTableWithWeights(FSMultipliersTable, Rand);
            Console.WriteLine($"FS multiplier = {fsMultiplier}"); // показываем множитель в отладочном режиме
            Console.ReadKey();
            // итоговый выигрыш спина = выигрыш по линиям * множитель фриспина
            win = paylinesWin * fsMultiplier;
            ClearMultiplierMatrixOutOfSymbolsMulties();
        }

        public void SetInnersSymbolsDuringFS()
        {
            InnerSymbolAtCurrentSpin = TableWeightsSelector.GetRandomObjectFromTableWithWeights(InnerReelSetsFS[CurrentFSReels], Rand);
        }

        public void RealizeInnerSymbols()
        {
            for (int i = 0; i < SlotWidth; i++)
            {
                for (int j = 0; j < SlotHeight; j++)
                {
                    if (_matrix[j, i] == Inner)
                    {
                        _matrix[j, i] = InnerSymbolAtCurrentSpin;
                    }
                }
            }
        }

        public void RealizeMultiplierMatrix()
        {
            for (int i = 0; i < SlotHeight; i++)
            {
                for (int j = 0; j < SlotWidth; j++)
                {
                    if (_matrix[i, j] == Wild && _multiplierMatrix[i, j] == 1)
                    {
                        _multiplierMatrix[i, j] = TableWeightsSelector.GetRandomObjectFromTableWithWeights(WildsMultipliersOfFS, Rand);
                    }
                    else if (_matrix[i, j] == _activatedSymbol)
                    {
                        _multiplierMatrix[i, j] = _currentMultiplierOfChosenSymbol;
                        _positionsOfMultiplierMatrixNeedsToBeClearedAfterSpin.Add((i, j));
                    }
                }
            }
        }

        public void ChooseCurrentFSReelSet()
        {
            CurrentFSReels = TableWeightsSelector.GetRandomObjectFromTableWithWeights(FSReelsWeights, Rand);
        }

        public void GenerateNewMatrix(Dictionary<int, List<Symbol>> reelSet)
        {
            GenerateNewStopPositions(reelSet);
            for (int i = 0; i < SlotWidth; i++)
            {
                for (int j = 0; j < SlotHeight; j++)
                {
                   // if (_matrix[j, i] != Wild)
                    _matrix[j, i] = reelSet[i][(StopPositions[i] + j) % reelSet[i].Count];
                }
            }
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
                    if (_matrix[j, i] == symbol)
                    {
                        result++;
                    }

                }
            }
            return result;
        }

        public void ClearMultiplierMatrix()
        {
            for (int i = 0; i < SlotHeight; i++)
            {
                for (int j = 0; j < SlotWidth; j++)
                {
                    _multiplierMatrix[i, j] = 1;
                }
            }
        }

        public void ClearMultiplierMatrixOutOfSymbolsMulties()
        {
            foreach (var item in _positionsOfMultiplierMatrixNeedsToBeClearedAfterSpin)
            {
                _multiplierMatrix[item.Item1, item.Item2] = 1;
            }
            _positionsOfMultiplierMatrixNeedsToBeClearedAfterSpin.Clear();
        }

        public void ClearMatrix()
        {
            for (int i = 0; i < SlotHeight; i++)
            {
                for (int j = 0; j < SlotWidth; j++)
                {
                    _matrix[i, j] = Blank;
                }
            }
        }

        public void ShowMultiplierMatrix()
        {
            Console.WriteLine($"Multiplier Matrix: ");
            for (int i = 0; i < SlotHeight; i++)
            {
                for (int j = 0; j < SlotWidth; j++)
                {
                    Console.Write($"{_multiplierMatrix[i, j]}  ");
                }
                Console.WriteLine();
            }
        }
    }
}