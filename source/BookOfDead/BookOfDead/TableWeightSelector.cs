using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static DevilWest.SlotInfo;
using static DevilWest.SlotStats;
using static DevilWest.PaylinesCounter;
using System.Threading;

namespace DevilWest
{
    public static class TableWeightsSelector
    {
        public static T GetRandomObjectFromTableWithWeights<T>(Dictionary<T, int> weightsTable, ThreadLocal<Random> rand)
        {
            int sumOfWeights = 0;

            foreach (var item in weightsTable)
            {
                sumOfWeights += item.Value;
            }

            int randomNumber = rand.Value.Next(0, sumOfWeights);
            int currentSum = 0;

            foreach (var item in weightsTable)
            {
                currentSum += item.Value;
                if (randomNumber < currentSum)
                {
                    return item.Key;
                }
            }
            return default;
        }


        public static List<Dictionary<Symbol, int>> GetRandomObjectFromTableWithWeights(Dictionary<List<Dictionary<Symbol, int>>, int> weightsTable, Random rand)
        {
            int sumOfWeights = 0;

            foreach (var item in weightsTable)
            {
                sumOfWeights += item.Value;
            }

            int randomNumber = rand.Next(0, sumOfWeights);
            int currentSum = 0;

            foreach (var item in weightsTable)
            {
                currentSum += item.Value;
                if (randomNumber < currentSum)
                {
                    return item.Key;
                }
            }

            return default;
        }



        public static T GetRandomObjectFromTableWithWeights<T>(List<(T, int)> weightsTable, Random rand)
        {
            int sumOfWeights = 0;
            foreach (var item in weightsTable)
            {
                sumOfWeights += item.Item2;
            }
            int i = 0;
            int randomNumber = rand.Next(0, sumOfWeights);
            int sumOfWrongWeights = weightsTable[i].Item2;
            while (randomNumber >= sumOfWrongWeights)
            {
                i++;
                sumOfWrongWeights += weightsTable[i].Item2;
            }
            return weightsTable[i].Item1;
        }


        public static T GetRandomObjectFromTableWithWeights<T>(Dictionary<T, double> weightsTable, Random rand)
        {
            double sumOfWeights = 0;
            foreach (var item in weightsTable)
            {
                sumOfWeights += item.Value;
            }
            int i = 0;
            double randomNumber = rand.NextDouble() * sumOfWeights;
            double sumOfWrongWeights = weightsTable.ElementAt(i).Value;
            while (randomNumber > sumOfWrongWeights)
            {
                i++;
                sumOfWrongWeights += weightsTable.ElementAt(i).Value;
            }
            return weightsTable.ElementAt(i).Key;
        }
    }
}
