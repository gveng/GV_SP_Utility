using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Numerics;

namespace WpfApp
{
    public sealed class TouchstoneFileData
    {
        public TouchstoneFileData(
            string filePath,
            int portCount,
            string format,
            string frequencyUnit,
            IReadOnlyList<string> parameterNames,
            IReadOnlyList<TouchstoneDataPoint> points)
        {
            FilePath = filePath;
            PortCount = portCount;
            Format = format;
            FrequencyUnit = frequencyUnit;
            ParameterNames = new ReadOnlyCollection<string>(parameterNames.ToList());
            Points = new ReadOnlyCollection<TouchstoneDataPoint>(points.ToList());
        }

        public string FilePath { get; }
        public string FileName => Path.GetFileName(FilePath);
        public int PortCount { get; }
        public string Format { get; }
        public string FrequencyUnit { get; }
        public IReadOnlyList<string> ParameterNames { get; }
        public IReadOnlyList<TouchstoneDataPoint> Points { get; }
    }

    public sealed class TouchstoneDataPoint
    {
        public TouchstoneDataPoint(double frequencyHz, IList<TouchstoneParameterValue> parameters)
        {
            FrequencyHz = frequencyHz;
            Parameters = new ReadOnlyCollection<TouchstoneParameterValue>(parameters);
        }

        public double FrequencyHz { get; }
        public IReadOnlyList<TouchstoneParameterValue> Parameters { get; }
    }

    public sealed class TouchstoneParameterValue
    {
        public TouchstoneParameterValue(string name, double real, double imaginary, double magnitude, double phaseDegrees)
        {
            Name = name;
            Real = real;
            Imaginary = imaginary;
            Magnitude = magnitude;
            PhaseDegrees = phaseDegrees;
        }

        public string Name { get; }
        public double Real { get; }
        public double Imaginary { get; }
        public double Magnitude { get; }
        public double PhaseDegrees { get; }
    }

    public static class TouchstoneParser
    {
        private static readonly Dictionary<string, double> FrequencyMultipliers = new(StringComparer.OrdinalIgnoreCase)
        {
            ["HZ"] = 1,
            ["KHZ"] = 1e3,
            ["MHZ"] = 1e6,
            ["GHZ"] = 1e9
        };

        public static TouchstoneFileData Parse(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("File not found", filePath);
            }

            var lines = File.ReadAllLines(filePath);
            var optionLine = lines.FirstOrDefault(l => l.TrimStart().StartsWith("#", StringComparison.Ordinal));
            if (optionLine == null)
            {
                throw new InvalidDataException("Option line (# ...) missing.");
            }

            var optionTokens = optionLine.TrimStart('#', ' ', '\t').Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
            if (optionTokens.Length < 3)
            {
                throw new InvalidDataException("Invalid option line (# ...).");
            }

            var frequencyUnit = optionTokens[0].ToUpperInvariant();
            var parameterType = optionTokens[1].ToUpperInvariant();
            var dataFormat = optionTokens[2].ToUpperInvariant();

            if (parameterType != "S")
            {
                throw new NotSupportedException($"Parameter type {parameterType} not supported (only S).");
            }

            if (!FrequencyMultipliers.TryGetValue(frequencyUnit, out var freqMultiplier))
            {
                throw new NotSupportedException($"Frequency unit {frequencyUnit} not supported.");
            }

            var dataBlocks = CollectDataBlocks(lines, optionLine);
            if (dataBlocks.Count == 0)
            {
                throw new InvalidDataException("No data found in touchstone file.");
            }

            int? detectedPorts = null;
            var points = new List<TouchstoneDataPoint>();
            var parameterNames = new List<string>();

            foreach (var blockTokens in dataBlocks)
            {
                if (blockTokens.Count < 3)
                {
                    continue;
                }

                var frequency = ParseDouble(blockTokens[0]) * freqMultiplier;
                var valueTokens = blockTokens.Skip(1).ToList();

                if (valueTokens.Count % 2 != 0)
                {
                    throw new InvalidDataException("Inconsistent number of values (expected value pairs for parameters).");
                }

                if (detectedPorts == null)
                {
                    var parameterCount = valueTokens.Count / 2;
                    var portGuess = Math.Sqrt(parameterCount);
                    var portCount = (int)Math.Round(portGuess);
                    if (portCount * portCount != parameterCount)
                    {
                        throw new InvalidDataException("Unable to determine port count from data entries.");
                    }

                    detectedPorts = portCount;
                    parameterNames = BuildParameterNames(portCount).ToList();
                }

                var ports = detectedPorts!.Value;
                var expectedParameterCount = ports * ports;
                if (valueTokens.Count != expectedParameterCount * 2)
                {
                    throw new InvalidDataException("Number of values does not match expected parameter count.");
                }

                var parameters = new List<TouchstoneParameterValue>(expectedParameterCount);
                for (var i = 0; i < expectedParameterCount; i++)
                {
                    var first = ParseDouble(valueTokens[2 * i]);
                    var second = ParseDouble(valueTokens[2 * i + 1]);

                    Complex complex;
                    double magnitude;
                    double phaseDeg;
                    double real;
                    double imaginary;

                    if (dataFormat == "RI")
                    {
                        real = first;
                        imaginary = second;
                        complex = new Complex(real, imaginary);
                        magnitude = complex.Magnitude;
                        phaseDeg = Math.Atan2(imaginary, real) * 180.0 / Math.PI;
                    }
                    else if (dataFormat == "MA")
                    {
                        magnitude = first;
                        phaseDeg = second;
                        complex = Complex.FromPolarCoordinates(magnitude, phaseDeg * Math.PI / 180.0);
                        real = complex.Real;
                        imaginary = complex.Imaginary;
                    }
                    else
                    {
                        throw new NotSupportedException($"Data format {dataFormat} not supported (only RI or MA).");
                    }

                    var name = parameterNames[i];
                    parameters.Add(new TouchstoneParameterValue(name, real, imaginary, magnitude, phaseDeg));
                }

                points.Add(new TouchstoneDataPoint(frequency, parameters));
            }

            return new TouchstoneFileData(filePath, detectedPorts ?? 0, dataFormat, frequencyUnit, parameterNames, points);
        }

        private static List<List<string>> CollectDataBlocks(string[] lines, string optionLine)
        {
            var blocks = new List<List<string>>();
            var current = new List<string>();
            var optionIndex = Array.IndexOf(lines, optionLine);
            for (var i = optionIndex + 1; i < lines.Length; i++)
            {
                var line = lines[i];
                if (string.IsNullOrWhiteSpace(line) || line.TrimStart().StartsWith("!", StringComparison.Ordinal))
                {
                    continue;
                }

                if (!char.IsWhiteSpace(line, 0))
                {
                    if (current.Count > 0)
                    {
                        blocks.Add(current);
                        current = new List<string>();
                    }
                }

                var tokens = line.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
                current.AddRange(tokens);
            }

            if (current.Count > 0)
            {
                blocks.Add(current);
            }

            return blocks;
        }

        private static IEnumerable<string> BuildParameterNames(int ports)
        {
            for (var j = 1; j <= ports; j++)
            {
                for (var i = 1; i <= ports; i++)
                {
                    yield return $"S{i}{j}";
                }
            }
        }

        private static double ParseDouble(string token)
        {
            return double.Parse(token, NumberStyles.Float | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
        }
    }
}
