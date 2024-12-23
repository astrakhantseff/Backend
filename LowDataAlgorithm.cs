using System;
using System.Linq;

namespace Prisma.Regressions.Math
{
    public class LowDataAlgorithm
    {
        private double[] _k;

        private InputParams _inputParams;

        private readonly InputParamsForPrognosis _inputParamsForPrognosis;

        public LowDataAlgorithm(InputParamsForPrognosis inputParamsForPrognosis, InputParams inputParams)
        {
            _inputParamsForPrognosis = inputParamsForPrognosis;
            _inputParams = inputParams;
        }

        public bool Step1()
        {
            bool result = FillAverageDebitStage();
            result &= FillDays();
            return result;
        }

        private bool FillAverageDebitStage()
        {
            try
            {
                double init = _inputParams.AverageDebit;
                double declineRate = _inputParams.DeclineRate;

                _inputParamsForPrognosis.qi[0] = init;
                for (int i = 1; i < _inputParamsForPrognosis.qi.Length; i++)
                {
                    // Вычисляем дебит как геометрическую прогрессию
                    _inputParamsForPrognosis.qi[i] = init * System.Math.Pow(1 - declineRate / 100, i);
                }

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка в вычислении прогрессии дебета {ex}");
            }
        }

        private bool FillDays()
        {
            try
            {
                for (int i = 0; i < _inputParamsForPrognosis.Ti.Length; i++)
                {
                    _inputParamsForPrognosis.Ti[i] = _inputParams.MinWorkingDays;
                }

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка ({ex})");
            }
        }

        public bool Step2(double[] rate)
        {
            try
            {
                //подбираем qi (также как дни, из лесенки)

                var length = _inputParamsForPrognosis.Years.Length;
                var result = new double[length, length];

                var wells = _inputParamsForPrognosis.Ni;
                var averageDebitYear = _inputParamsForPrognosis.qi;
                var averageDays = _inputParamsForPrognosis.Ti;
                _k = GetK(rate);

                double value;

                for (int col = 0; col < length; col++)
                {
                    for (int row = 0; row < length; row++)
                    {
                        if (row < col)
                        {
                            if (row == 0)
                                value = wells[row] * _k[col] * averageDebitYear[row];
                            else
                                value = wells[row] * _k[col - row] * averageDebitYear[row];

                            result[row, col] = value;
                        }
                        else if (row == col)
                        {
                            value = wells[row] * _k[0] * averageDebitYear[row] * averageDays[row] / 365;

                            result[row, col] = value;
                            break;
                        }
                    }
                }
                _inputParamsForPrognosis.StepWellDays = result;

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка ({ex})");
            }
        }

        public static double[] GetSum(double[,] rate)
        {
            // Создаем массив для хранения результатов
            double[] result = new double[rate.GetLength(1)];

            // Используем LINQ для вычисления сумм по столбцам
            result = Enumerable.Range(0, rate.GetLength(1))
                               .Select(col => Enumerable.Range(0, rate.GetLength(0))
                                                        .Select(row => rate[row, col])
                                                        .Sum())
                               .ToArray();

            return result;
        }

        private double[] GetK(double[] rate)
        {
            var result = rate.Select(r => r / 1000 * (365 * _inputParams.Ke)).ToArray();
            return result;
        }
    }
}