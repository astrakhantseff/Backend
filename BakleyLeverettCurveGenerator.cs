using System.Collections.Generic;

namespace Prisma.Regressions.Math
{
    public class BakleyLeverettCurveGenerator
    {
        // Параметры задачи
        public double Swc { get; set; } = 0.0; // остаточная вода насыщение
        public double Sor { get; set; } = 0.0; // остаточное масло насыщение
        public double ViscosityRatio { get; set; } = 0.5; // отношение вязкости воды к вязкости нефти

        // Метод для генерации одной кривой в зависимости от начальной насыщенности воды
        public List<(double, double)> GenerateSingleCurve()
        {
            return GetBuckleyLeverettData(Swc, Sor, ViscosityRatio);
        }

        // Функция для вычисления Saturation (S) и Fractional Flow (F)
        private static List<(double, double)> GetBuckleyLeverettData(double swc, double sor, double viscosityRatio)
        {
            var data = new List<(double, double)>();

            double deltaSw = (1 - swc - sor) / Constants.NumberOfWcCalc;

            for (int i = 0; i <= Constants.NumberOfWcCalc; i++)
            {
                double sw = swc + i * deltaSw;
                double fw = FractionalFlow(sw, swc, sor, viscosityRatio);
                data.Add((sw, fw));
            }

            return data;
        }

        // Функция для вычисления фракционного расхода (fractional flow)
        private static double FractionalFlow(double sw, double swc, double sor, double viscosityRatio)
        {
            // Простейшая модель для фракционного расхода
            double krw = System.Math.Pow((sw - swc) / (1 - swc - sor), 2); // относительная проницаемость воды
            double kro = System.Math.Pow((1 - sw - sor) / (1 - swc - sor), 2); // относительная проницаемость нефти

            return krw / (krw + viscosityRatio * kro);
        }
    }
}
