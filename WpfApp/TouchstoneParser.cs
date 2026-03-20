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
            double referenceImpedance,
            IReadOnlyList<string> parameterNames,
            IReadOnlyList<TouchstoneDataPoint> points)
        {
            FilePath = filePath;
            PortCount = portCount;
            Format = format;
            FrequencyUnit = frequencyUnit;
            ReferenceImpedance = referenceImpedance;
            ParameterNames = new ReadOnlyCollection<string>(parameterNames.ToList());
            Points = new ReadOnlyCollection<TouchstoneDataPoint>(points.ToList());
        }

        public string FilePath { get; }
        public string FileName => Path.GetFileName(FilePath);
        public int PortCount { get; }
        public string Format { get; }
        public string FrequencyUnit { get; }
        public double ReferenceImpedance { get; }
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
            double referenceImpedance = 50.0;

            // Search for R and its value
            for (var i = 3; i < optionTokens.Length; i++)
            {
                if (string.Equals(optionTokens[i], "R", StringComparison.OrdinalIgnoreCase) && i + 1 < optionTokens.Length)
                {
                    if (double.TryParse(optionTokens[i + 1], NumberStyles.Float, CultureInfo.InvariantCulture, out var r))
                    {
                        referenceImpedance = r;
                        break;
                    }
                }
            }

            if (parameterType != "S")
            {
                throw new NotSupportedException($"Parameter type {parameterType} not supported (only S).");
            }

            if (!FrequencyMultipliers.TryGetValue(frequencyUnit, out var freqMultiplier))
            {
                throw new NotSupportedException($"Frequency unit {frequencyUnit} not supported.");
            }

            // First pass: detect number of ports from # line or extension
            int? detectedPorts = null;
            // Extension check: .sNb where N is ports
            var ext = Path.GetExtension(filePath)?.ToLowerInvariant();
            if (!string.IsNullOrEmpty(ext) && ext.Length >= 3 && ext.StartsWith(".s") && ext.EndsWith("p"))
            {
                var numPart = ext.Substring(2, ext.Length - 3);
                if (int.TryParse(numPart, out int p))
                {
                    detectedPorts = p;
                }
            }

            // Collect data blocks knowing expected ports if possible
            var dataBlocks = CollectDataBlocks(lines, optionLine, detectedPorts);
            if (dataBlocks.Count == 0)
            {
                throw new InvalidDataException("No data found in touchstone file.");
            }

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

                // Logic to align number of data points
                if (detectedPorts == null)
                {
                    // Heuristic fallback if extension didn't help
                    var parameterCount = valueTokens.Count / 2;
                    var portGuess = Math.Sqrt(parameterCount);
                    var portCount = (int)Math.Round(portGuess);
                    
                    if (portCount * portCount != parameterCount || portCount <= 0)
                    {
                         // If we can't guess, we can't proceed safely
                         throw new InvalidDataException("Unable to determine port count from data entries.");
                    }

                    detectedPorts = portCount;
                }
                
                if (parameterNames.Count == 0)
                {
                     parameterNames = BuildParameterNames(detectedPorts.Value).ToList();
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

            return new TouchstoneFileData(filePath, detectedPorts ?? 0, dataFormat, frequencyUnit, referenceImpedance, parameterNames, points);
        }

        private static List<List<string>> CollectDataBlocks(string[] lines, string optionLine, int? expectedPorts)
        {
            var blocks = new List<List<string>>();
            var currentTokens = new List<string>();
            
            // Find option line index to start scanning after it
            var startLineIndex = 0;
            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i].Trim();
                if (line.StartsWith("#"))
                {
                    startLineIndex = i + 1;
                    break;
                }
            }

            for (var i = startLineIndex; i < lines.Length; i++)
            {
                var line = lines[i];
                if (string.IsNullOrWhiteSpace(line) || line.TrimStart().StartsWith("!", StringComparison.Ordinal))
                {
                    continue;
                }

                var tokens = line.Split((char[])null!, StringSplitOptions.RemoveEmptyEntries);
                if (tokens.Length == 0) continue;

                // Check if this line starts a new frequency block
                // A new block starts with a number that is the frequency.
                // However, continuation lines also contain numbers.
                // Standard Touchstone files typically have the frequency as the first token of a new block.
                // The most robust way for N-port (where N is known) is to count tokens.
                
                if (expectedPorts.HasValue)
                {
                    // If we know the number of ports (e.g. from extension), we ignore line breaks/indentation logic
                    // and just collect everything into a single token stream, then chunk it later.
                    // This creates one giant "currentTokens" list which is added to blocks at the end.
                    currentTokens.AddRange(tokens);
                }
                else
                {
                    // Fallback to indentation heuristic if ports are unknown
                    // If line starts with non-whitespace, it's a new block (frequency)
                    if (!char.IsWhiteSpace(line, 0))
                    {
                        if (currentTokens.Count > 0)
                        {
                            blocks.Add(currentTokens);
                            currentTokens = new List<string>();
                        }
                    }
                    currentTokens.AddRange(tokens);
                }
            }

            if (currentTokens.Count > 0)
            {
                blocks.Add(currentTokens);
            }
            
            // If we have known ports, flatten all tokens and chunk them into correct block sizes.
            if (expectedPorts.HasValue)
            {
                 return RegroupBlocksByCount(blocks, expectedPorts.Value);
            }

            return blocks;
        }

        private static List<List<string>> RegroupBlocksByCount(List<List<string>> rawBlocks, int textPorts)
        {
             var allTokens = new List<string>();
             foreach(var b in rawBlocks) allTokens.AddRange(b);
             
             var result = new List<List<string>>();
             // Each block is: Frequency + (N * N * 2 parameters)
             int tokensPerBlock = 1 + 2 * textPorts * textPorts;
             
             if (tokensPerBlock <= 0) return result; // safety
             
             int index = 0;
             while(index + tokensPerBlock <= allTokens.Count)
             {
                 var chunk = allTokens.GetRange(index, tokensPerBlock);
                 result.Add(chunk);
                 index += tokensPerBlock;
             }
             return result;
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
