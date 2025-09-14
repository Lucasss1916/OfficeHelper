using System;
using System.Collections.Generic;
using System.Linq;

namespace RandomScoreAllocatorWPF
{
    public static class ScoreAllocator
    {
        /// <summary>
        /// 将 totalScore 随机分配到 subjects 上，带权重与边界，输出整数且总和严格等于 totalScore。
        /// </summary>
        public static Dictionary<string, int> Allocate(
            int totalScore,
            IList<string> subjects,
            IList<double>? weights = null,
            int minEach = 1,
            double maxEachFraction = 0.6,
            double randomness = 0.25,
            int? seed = null)
        {
            if (subjects == null || subjects.Count == 0)
                throw new ArgumentException("subjects 不能为空。");
            int n = subjects.Count;
            if (totalScore < n * minEach)
                throw new ArgumentException($"totalScore={totalScore} 太小，无法满足 n*minEach={n * minEach}。");

            // 权重
            double[] w = (weights == null || weights.Count != n)
                ? Enumerable.Repeat(1.0, n).ToArray()
                : weights.Select(x => Math.Max(0.0001, x)).ToArray();
            double wSum = w.Sum();
            double[] p = w.Select(x => x / wSum).ToArray();

            // 期望目标
            double[] targets = p.Select(pi => totalScore * pi).ToArray();

            int maxEach = (int)Math.Floor(totalScore * maxEachFraction);
            maxEach = Math.Max(maxEach, minEach + 1);

            var rnd = seed.HasValue ? new Random(seed.Value) : new Random();
            double sigma = randomness * totalScore / Math.Sqrt(n);

            double[] noisy = new double[n];
            for (int i = 0; i < n; i++)
            {
                double noise = NextGaussian(rnd, 0, sigma);
                noisy[i] = Math.Max(minEach, Math.Min(maxEach, targets[i] + noise));
            }

            int[] ints = noisy.Select(x => (int)Math.Round(x)).ToArray();
            int diff = totalScore - ints.Sum();
            if (diff != 0)
            {
                AdjustToSum(ints, diff, minEach, maxEach, rnd);
            }

            // 小概率全等时打散
            if (ints.All(x => x == ints[0]) && n >= 2)
            {
                int a = rnd.Next(n);
                int b = (a + 1) % n;
                if (ints[a] > minEach && ints[b] < maxEach)
                {
                    ints[a] -= 1;
                    ints[b] += 1;
                }
            }

            var result = new Dictionary<string, int>(n);
            for (int i = 0; i < n; i++) result[subjects[i]] = ints[i];
            if (result.Values.Sum() != totalScore)
                throw new InvalidOperationException("分配失败：和不等于总分。");
            return result;
        }

        private static double NextGaussian(Random rnd, double mean, double stddev)
        {
            double u1 = 1.0 - rnd.NextDouble();
            double u2 = 1.0 - rnd.NextDouble();
            double z0 = Math.Sqrt(-2.0 * Math.Log(u1)) * Math.Cos(2 * Math.PI * u2);
            return mean + z0 * stddev;
        }

        private static void AdjustToSum(int[] ints, int diff, int minEach, int maxEach, Random rnd)
        {
            int n = ints.Length;
            int steps = Math.Abs(diff);
            int sign = diff > 0 ? 1 : -1;
            var idx = Enumerable.Range(0, n).ToArray();

            for (int s = 0; s < steps; s++)
            {
                bool done = false;
                for (int t = 0; t < 5 * n; t++)
                {
                    int i = idx[rnd.Next(n)];
                    if (sign > 0 && ints[i] < maxEach) { ints[i]++; done = true; break; }
                    if (sign < 0 && ints[i] > minEach) { ints[i]--; done = true; break; }
                }
                if (!done)
                {
                    if (sign > 0)
                    {
                        int i = Array.FindIndex(ints, x => x < maxEach);
                        if (i >= 0) { ints[i]++; done = true; }
                    }
                    else
                    {
                        int i = Array.FindIndex(ints, x => x > minEach);
                        if (i >= 0) { ints[i]--; done = true; }
                    }
                }
                if (!done) break;
            }
        }
    }
}
