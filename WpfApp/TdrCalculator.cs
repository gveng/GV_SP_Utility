using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;

namespace WpfApp
{
    public enum WindowType
    {
        Rectangular,
        Hamming,
        Hanning,
        Blackman
    }

    public class TdrSettings
    {
        public double RiseTime { get; set; } = 30e-12; // 30ps default
        public WindowType Window { get; set; } = WindowType.Hanning;
        public double Delay { get; set; } = 0;
        public double SystemImpedance { get; set; } = 50.0;
        public double MaxDuration { get; set; } = 10e-9;
    }

    public class TdrResult
    {
        public double[] Time { get; set; }
        public double[] Impedance { get; set; }
        public double[] Reflection { get; set; } // Step response
    }

    public static class TdrCalculator
    {
        public static TdrResult Calculate(TouchstoneFileData file, string paramName, TdrSettings settings)
        {
            // 1. Extract data and find max frequency
            var points = file.Points;
            if (points == null || points.Count == 0)
            {
                throw new InvalidOperationException("File contains no data points.");
            }

            double maxFreq = points.Last().FrequencyHz;
            
            // Calculate required points to cover the desired Time Duration
            // T_max = 1 / deltaF => deltaF = 1 / T_max
            // Max Points N = 2 * Fmax / deltaF = 2 * Fmax * T_max
            
            double targetTime = settings.MaxDuration;
            if(targetTime <= 0) targetTime = 10e-9; // Fallback
            
            double requiredPointsDouble = 2 * maxFreq * targetTime;
            
            // Limit nPoints to avoid excessive memory usage/calculation time
            // Typical max TDR points: 16k or 64k is plenty.
            int nPoints = NextPowerOfTwo((int)requiredPointsDouble); 
            if (nPoints < 1024) nPoints = 1024;
            if (nPoints > 131072) nPoints = 131072; // Increased limit slightly to allow longer duration if needed (approx 131k points is fast enough on modern CPU)

            double deltaF = maxFreq / (nPoints / 2); // Frequency step from DC to MaxFreq

            // 2. Resample to linear frequency (including DC extrapolation)
            Complex[] sParamFreq = ResampleAndExtrapolate(points, paramName, nPoints, deltaF);
            
            // Check for NaN/Inf
            if (sParamFreq.Any(c => double.IsNaN(c.Real) || double.IsNaN(c.Imaginary)))
            {
                throw new ArithmeticException("Numeric error during interpolation (NaN detected).");
            }

            // 3. Apply Filtering (Gaussian Rise Time)
            ApplyRiseTimeFilter(sParamFreq, settings.RiseTime, deltaF);

            // 4. Apply Windowing
            ApplyWindow(sParamFreq, settings.Window);

            // Apply Phase Delay (Frequency Domain Convolution/Shift)
            // Ideally we want the Step to occur at t=Delay.
            // S_new(f) = S(f) * exp( -j * 2 * pi * f * Delay )
            // Be careful about the sign.
            // Delay > 0 means the response is delayed (right shift).
            if (settings.Delay != 0)
            {
                 ApplyDelay(sParamFreq, settings.Delay, deltaF);
            }

            // 5. IFFT to get Impulse Response
            Complex[] timeDomainComplex = ComputeIFFT(sParamFreq);

            double[] impulseResponse = timeDomainComplex.Select(c => c.Real).ToArray();

            // 6. Integrate to get Step Response (Reflection Coefficient Rho(t))
            double[] stepResponse = IntegrateImpulse(impulseResponse);

            // 7. Calculate Impedance Z(t)
            double[] impedance = CalculateImpedanceProfile(stepResponse, settings.SystemImpedance);

            // 8. Generate Time axis
            double totalTime = 1.0 / deltaF;
            double dt = totalTime / nPoints;

            double[] timeAxis = new double[nPoints];
            for (int i = 0; i < nPoints; i++)
            {
                // Time starts at 0, Delay is embedded in the signal via phase shift
                timeAxis[i] = i * dt; 
            }

            return new TdrResult
            {
                Time = timeAxis,
                Impedance = impedance,
                Reflection = stepResponse
            };
        }

        private static int NextPowerOfTwo(int n)
        {
            int p = 1;
            while (p < n) p <<= 1;
            return p;
        }

        private static Complex[] ResampleAndExtrapolate(IReadOnlyList<TouchstoneDataPoint> points, string paramName, int nPoints, double deltaF)
        {
            // nPoints is the full IFFT size. We only fill 0 to N/2 (Positive frequencies).
            // Index 0 is DC. Index N/2 is Nyquist (MaxFreq).
            int numPositivePoints = nPoints / 2 + 1;
            Complex[] data = new Complex[nPoints]; // Initialize with zeros

            // Extract original data for interpolation
            var originalData = new List<(double f, Complex val)>();
            foreach (var p in points)
            {
                var param = p.Parameters.FirstOrDefault(x => x.Name == paramName);
                if (param != null)
                {
                    originalData.Add((p.FrequencyHz, new Complex(param.Real, param.Imaginary)));
                }
            }

            if (originalData.Count == 0)
            {
                // No data found for parameter, return zeros (avoid crash)
                return data;
            }

            // Optimize sequential access by tracking index
            int lastIndexObj = 0;

            // Interpolate
            for (int i = 0; i < numPositivePoints; i++)
            {
                double targetFreq = i * deltaF;

                if (i == 0)
                {
                    // Extrapolate DC
                    // Simple strategy: Use the magnitude of the first point, phase 0 (resistive) or linear extrapolation?
                    // Safe bet: Duplicate first point's real part, set imag to 0 (assuming real impedance at DC)
                    // Or linear extrapolation from first two points.
                    if (originalData.Count >= 2)
                    {
                        var p1 = originalData[0];
                        var p2 = originalData[1];
                        // Linear extrapolation for Real and Imag
                        double re = LinearExtrap(p1.f, p1.val.Real, p2.f, p2.val.Real, 0);
                        // Imaginary part usually 0 at DC for physical passive systems?
                        // Let's stick with 0 imaginary at DC for robustness.
                        data[i] = new Complex(re, 0); 
                    }
                    else if (originalData.Count == 1)
                    {
                         data[i] = new Complex(originalData[0].val.Real, 0);
                    }
                    else
                    {
                        data[i] = Complex.Zero;
                    }
                }
                else
                {
                    // Find surrounding points using optimized search
                    // Pass lastIndexObj by ref to update it
                    data[i] = Interpolate(originalData, targetFreq, ref lastIndexObj);
                }
            }

            // Create Hermitian symmetry for real IFFT result
            // F[N-i] = Conj(F[i])
            for (int i = 1; i < nPoints / 2; i++)
            {
                data[nPoints - i] = Complex.Conjugate(data[i]);
            }
            // Nyquist point (N/2) must be real
            data[nPoints / 2] = new Complex(data[nPoints / 2].Real, 0);

            return data;
        }

        private static double LinearExtrap(double x1, double y1, double x2, double y2, double targetX)
        {
             if (Math.Abs(x2 - x1) < 1e-9) return y1; // Avoid divide by zero
             double slope = (y2 - y1) / (x2 - x1);
             return y1 + slope * (targetX - x1);
        }

        private static Complex Interpolate(List<(double f, Complex val)> data, double targetFreq, ref int lastIndex)
        {
            // Optimized interpolation assuming sequential access and sorted data
            
            // Clamp
            if (targetFreq >= data[data.Count - 1].f) return data[data.Count - 1].val;
            if (targetFreq <= data[0].f) return data[0].val;

            // Fast forward from lastIndex
            // Use local variable to avoid excessive property access in debug
            int count = data.Count;
            int idx = lastIndex;

            // Usually we move forward by 0 or 1 step if deltaF is small
            while (idx < count - 1 && data[idx + 1].f < targetFreq)
            {
                idx++;
            }
            // In case targetFreq is smaller (backward search - unlikely in this loop order but safe)
            while (idx > 0 && data[idx].f > targetFreq)
            {
                idx--;
            }

            lastIndex = idx; // Update hint

            // Now data[idx].f <= targetFreq < data[idx+1].f
            // Or roughly there.
            
            // Boundary safety
            if (idx >= count - 1) return data[count - 1].val;
            
            var pLeft = data[idx];
            var pRight = data[idx + 1];

            // Re-check just in case logic above slipped
            if (targetFreq < pLeft.f) return pLeft.val; 
            
            double divisor = pRight.f - pLeft.f;
            if (divisor < 1e-9) return pLeft.val; // Avoid div/0

            double t = (targetFreq - pLeft.f) / divisor;
            
            double re = pLeft.val.Real + t * (pRight.val.Real - pLeft.val.Real);
            double im = pLeft.val.Imaginary + t * (pRight.val.Imaginary - pLeft.val.Imaginary);

            return new Complex(re, im);
        }

        private static void ApplyRiseTimeFilter(Complex[] data, double riseTime, double deltaF)
        {
            // Gaussian Filter
            // H(f) = exp( - (f / f_c)^2 * ln(2) ) ? No, standard Gaussian rise time relate
            // BW (3dB) = 0.35 / Tr
            // Sigma relation?
            // H(f) = exp( - (2 * pi * f * sigma)^2 / 2 )
            // Tr (10-90%) approx 2.56 * sigma? Or 3.3?
            // Standard approx: f_3dB = 0.35 / Tr
            // H(f) = 1 / sqrt(1 + (f/f_3dB)^2) is 1st order. Gaussian is exp(-0.3466 * (f * Tr)^2) ?
            
            // Let's use the formula: Gamma(t) convolved with Gaussian.
            // Frequency domain: S(f) * G(f)
            // G(f) = exp( - ( f * Tr * (Pi / sqrt(ln(2)) ? ) ) )
            
            // Reference: keysight or similar
            // Gamma = alpha * f^2
            // alpha = (tr / (4 * erfinv(0.8)))^2 ... too complex without math lib
            
            // Lets use: sigma = Tr / 1.665 (approx for 10-90 on Gaussian step) or similar
            // Let's use standard approximation:
            // sigma_f = 0.35 / Tr (BW)
            // Gaussian: exp ( - f^2 / (2 * sigma_f^2) ) ... no that's bell curve.
            
            // Correct formula often used in TDR sims:
            // H(f) = exp( - ((f * Tr) / k)^2 )
            // where k is constant.
            
            // Let's use f_3dB = 0.35 / Tr.
            // H(f) = exp( - ln(2) * (f / f_3dB)^2 )
            
            if (riseTime <= 0) return;

            double f_3dB = 0.35 / riseTime;
            int n = data.Length;
            int half = n / 2;

            for (int i = 0; i <= half; i++)
            {
                double f = i * deltaF;
                double magnitudeFactor = Math.Exp(-Math.Log(2) * Math.Pow(f / f_3dB, 2));
                
                data[i] *= magnitudeFactor;
                if (i > 0 && i < half)
                {
                    data[n - i] *= magnitudeFactor; // Apply to negative frequencies
                }
            }
        }

        private static void ApplyDelay(Complex[] data, double delay, double deltaF)
        {
            if (delay == 0) return;
            // Delay in seconds -> Phase shift
            int n = data.Length;
            int half = n / 2;
            
            for (int i = 0; i <= half; i++)
            {
                // angle = -2 * pi * f * t
                double f = i * deltaF;
                double angle = -2 * Math.PI * f * delay;
                
                Complex phaseShift = Complex.FromPolarCoordinates(1, angle);
                data[i] *= phaseShift;
                
                if (i > 0 && i < half)
                {
                    data[n - i] = Complex.Conjugate(data[i]);
                }
            }
        }

        private static void ApplyWindow(Complex[] data, WindowType type)
        {
            if (type == WindowType.Rectangular) return;

            int n = data.Length;
            int half = n / 2;
            
            // Window is applied to the frequency domain to reduce ringing (Gibbs phenomenon)
            // Usually applied from DC to Fmax (0 to half)
            
            for (int i = 0; i <= half; i++)
            {
                double w = 1.0;
                double ratio = (double)i / half; // 0 to 1

                switch (type)
                {
                    case WindowType.Hamming:
                        w = 0.54 + 0.46 * Math.Cos(Math.PI * ratio);
                        break;
                    case WindowType.Hanning:
                        w = 0.5 * (1 + Math.Cos(Math.PI * ratio));
                        break;
                    case WindowType.Blackman:
                        w = 0.42 + 0.5 * Math.Cos(Math.PI * ratio) + 0.08 * Math.Cos(2 * Math.PI * ratio);
                        break;
                }

                data[i] *= w;
                if (i > 0 && i < half)
                {
                   data[n - i] *= w;
                }
            }
        }

        private static Complex[] ComputeIFFT(Complex[] input)
        {
            // Basic Cooley-Tukey FFT implementation
            // Since we need IFFT, we can use FFT and swap / scale
            // IFFT(x) = conj(FFT(conj(x))) / N
            
            int n = input.Length;
            Complex[] conjInput = new Complex[n];
            for(int i=0; i<n; i++) conjInput[i] = Complex.Conjugate(input[i]);
            
            Complex[] fftResult = FFT(conjInput);
            
            Complex[] result = new Complex[n];
            for(int i=0; i<n; i++)
            {
                result[i] = Complex.Conjugate(fftResult[i]) / n;
            }
            return result;
        }

        private static Complex[] FFT(Complex[] x)
        {
            int n = x.Length;
            if (n == 1) return new Complex[] { x[0] };

            if (n % 2 != 0) throw new ArgumentException("N must be power of 2");

            Complex[] even = new Complex[n / 2];
            Complex[] odd = new Complex[n / 2];
            
            for (int i = 0; i < n / 2; i++)
            {
                even[i] = x[2 * i];
                odd[i] = x[2 * i + 1];
            }

            Complex[] e = FFT(even);
            Complex[] o = FFT(odd);
            
            Complex[] result = new Complex[n];
            for (int k = 0; k < n / 2; k++)
            {
                double angle = -2 * Math.PI * k / n;
                Complex w = Complex.FromPolarCoordinates(1, angle);
                result[k] = e[k] + w * o[k];
                result[k + n / 2] = e[k] - w * o[k];
            }
            return result;
        }

        private static double[] IntegrateImpulse(double[] impulse)
        {
            double[] step = new double[impulse.Length];
            double sum = 0;
            // The impulse response from IFFT of S11 is the reflection coefficient impulse.
            // The step response is the integral.
            // Scaling is handled by the DC term in frequency domain usually, 
            // but IFFT needs to be careful.
            // If DC component S11(0) was correct, the step response should settle to S11(0).
            // But we often look at TDR as Impedance.
            
            for (int i = 0; i < impulse.Length; i++)
            {
                sum += impulse[i];
                step[i] = sum;
            }
            
            return step;
        }

        private static double[] CalculateImpedanceProfile(double[] rho, double z0)
        {
            double[] z = new double[rho.Length];
            for (int i = 0; i < rho.Length; i++)
            {
                // Z = Z0 * (1 + rho) / (1 - rho)
                // rho is the reflection coefficient at this point in time (cumulative)
                
                double r = rho[i];
                // Clamp r to avoid division by zero or negative impedance extremes
                if (r > 0.999) r = 0.999;
                if (r < -0.999) r = -0.999;
                
                z[i] = z0 * (1 + r) / (1 - r);
            }
            return z;
        }
    }
}
