using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;

namespace WpfApp
{
    public class RlcResult
    {
        public double R { get; set; }
        public double L { get; set; }
        public double C { get; set; }
        public double ResonanceFreq { get; set; }
    }

    public static class RlcHelper
    {
        public static Complex CalculateComplexImpedance(TouchstoneParameterValue param, double systemImpedance)
        {
            var complexParam = new Complex(param.Real, param.Imaginary);

            bool isTransmission = param.Name.Contains("S21", StringComparison.OrdinalIgnoreCase) ||
                                  param.Name.Contains("S12", StringComparison.OrdinalIgnoreCase);

            if (isTransmission)
            {
                // Series-Thru Method (same as GraphWindow)
                // Z = 2 * Z0 * (1 - S21) / S21
                if (complexParam.Magnitude < 1e-12) return new Complex(double.PositiveInfinity, double.PositiveInfinity);
                return 2 * systemImpedance * (1 - complexParam) / complexParam;
            }
            else
            {
                // Reflection Method (1-port Z)
                // Z = Z0 * (1 + S11) / (1 - S11)
                if ((1 - complexParam).Magnitude < 1e-12) return new Complex(double.PositiveInfinity, double.PositiveInfinity);
                return systemImpedance * (1 + complexParam) / (1 - complexParam);
            }
        }

        public static double CalculateComplexImpedanceMagnitude(TouchstoneParameterValue param, double systemImpedance)
        {
            var z = CalculateComplexImpedance(param, systemImpedance);
            return z.Magnitude;
        }

        public static RlcResult FitSeriesRlc(List<(double Freq, Complex Z)> data)
        {
            if (data == null || data.Count == 0) return new RlcResult();

            var minZ = data.OrderBy(d => d.Z.Magnitude).First();
            double fRes = minZ.Freq;
            double R = minZ.Z.Real;

            double wRes = 2 * Math.PI * fRes;
            double LC_Product = 1.0 / (wRes * wRes);

            var cPoints = data.Where(d => d.Freq < fRes && d.Z.Imaginary < 0).ToList();
            double C_est = 0;
            if (cPoints.Count > 0)
            {
                double sumC = 0;
                int validCount = 0;
                foreach (var p in cPoints)
                {
                    double w = 2 * Math.PI * p.Freq;
                    double val = -1.0 / (w * p.Z.Imaginary);
                    if (val > 0 && val < 1)
                    {
                        sumC += val;
                        validCount++;
                    }
                }
                if (validCount > 0) C_est = sumC / validCount;
            }

            double L_calc = 0;
            if (C_est > 0)
            {
                L_calc = LC_Product / C_est;
            }

            if (C_est == 0)
            {
                var lPoints = data.Where(d => d.Freq > fRes && d.Z.Imaginary > 0).ToList();
                double L_est = 0;
                if (lPoints.Count > 0)
                {
                    double sumL = 0;
                    int validCount = 0;
                    foreach (var p in lPoints)
                    {
                        double w = 2 * Math.PI * p.Freq;
                        double val = p.Z.Imaginary / w;
                        if (val > 0)
                        {
                            sumL += val;
                            validCount++;
                        }
                    }
                    if (validCount > 0) L_est = sumL / validCount;
                }

                if (L_est > 0)
                {
                    L_calc = L_est;
                    C_est = LC_Product / L_calc;
                }
            }

            return new RlcResult { R = R, L = L_calc, C = C_est, ResonanceFreq = fRes };
        }

        public static string ToEngineeringNotation(double value, string unit)
        {
            if (value == 0) return $"0 {unit}";
            if (double.IsInfinity(value) || double.IsNaN(value)) return value.ToString();

            double mag = Math.Abs(value);
            int exponent = (int)Math.Floor(Math.Log10(mag));
            int engExponent = (exponent >= 0) ? (exponent / 3) * 3 : ((exponent - 2) / 3) * 3;

            double scaledValue = value / Math.Pow(10, engExponent);

            string prefix = engExponent switch
            {
                12 => "T",
                9 => "G",
                6 => "M",
                3 => "k",
                0 => "",
                -3 => "m",
                -6 => "µ",
                -9 => "n",
                -12 => "p",
                -15 => "f",
                _ => "?"
            };

            if (prefix == "?")
            {
                return $"{value:E3} {unit}";
            }

            return $"{scaledValue:F3} {prefix}{unit}";
        }
    }
}
