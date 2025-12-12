using System;
using System.Collections.Generic;
using System.Linq;

namespace RandomScoreAllocatorWPF
{
    public static class ScoreAllocator
    {
        public static Dictionary<string, int> Allocate(
            int totalScore,
            IList<string> subjects,
            IList<int> maxScores,
            int minEach = 0,
            double randomness = 0.25,
            int? seed = null)
        {
            if (subjects == null || subjects.Count == 0) return new Dictionary<string, int>();

            int n = subjects.Count;
            long paperMax = maxScores.Sum(x => (long)x); // 计算卷面总满分

            // 【关键校验 1】绝对不允许目标分超过卷面总满分
            // 如果 Excel 里写总分 100，但所有科目加起来只有 95，那只能分 95 分，否则必爆表
            if (totalScore > paperMax) totalScore = (int)paperMax;

            // 【关键校验 2】不能低于保底分
            if (totalScore < n * minEach) totalScore = n * minEach;

            // 1. 期望分配 (按比例)
            var rnd = seed.HasValue ? new Random(seed.Value) : new Random();
            double[] targets = new double[n];
            // 避免除以0
            double safePaperMax = paperMax == 0 ? 1 : paperMax;

            for (int i = 0; i < n; i++)
            {
                double ratio = (double)maxScores[i] / safePaperMax;
                targets[i] = totalScore * ratio;
            }

            // 2. 随机生成 + 初步钳制
            int[] ints = new int[n];
            double sigma = randomness * (totalScore / Math.Sqrt(n));

            for (int i = 0; i < n; i++)
            {
                double noise = NextGaussian(rnd, 0, sigma);
                int val = (int)Math.Round(targets[i] + noise);

                // 严格钳制：绝对不能超过 maxScores[i]
                ints[i] = Math.Max(minEach, Math.Min(maxScores[i], val));
            }

            // 3. 【核心逻辑】从后向前寻找空位补差
            AdjustToSum(ints, totalScore, minEach, maxScores);

            // 4. 输出
            var result = new Dictionary<string, int>();
            for (int i = 0; i < n; i++) result[subjects[i]] = ints[i];
            return result;
        }

        private static void AdjustToSum(int[] ints, int targetTotal, int minEach, IList<int> maxScores)
        {
            int currentSum = ints.Sum();
            if (currentSum == targetTotal) return;

            int n = ints.Length;
            // 安全阀：理论上只要 targetTotal <= paperMax，这个循环一定能结束
            int maxLoop = 100000;
            int loop = 0;

            while (currentSum != targetTotal)
            {
                loop++;
                if (loop > maxLoop) break; // 防止极端数据导致死锁

                bool adjusted = false;

                // --- 情况 A: 分少了 (Current < Target) -> 需要加分 ---
                if (currentSum < targetTotal)
                {
                    // 策略：从后向前遍历，寻找【未满分】的科目
                    for (int i = n - 1; i >= 0; i--)
                    {
                        if (ints[i] < maxScores[i]) // 只要这一科还没满
                        {
                            ints[i]++;      // 加 1 分
                            currentSum++;   // 更新总和
                            adjusted = true;
                            break;          // 加完这一次，退出 for，重新检查 while
                        }
                    }
                }
                // --- 情况 B: 分多了 (Current > Target) -> 需要减分 ---
                else
                {
                    // 策略：从后向前遍历，寻找【高于最低分】的科目
                    for (int i = n - 1; i >= 0; i--)
                    {
                        if (ints[i] > minEach) // 只要这一科还能减
                        {
                            ints[i]--;      // 减 1 分
                            currentSum--;   // 更新总和
                            adjusted = true;
                            break;          // 减完这一次，退出 for，重新检查 while
                        }
                    }
                }

                // 如果遍历了一整圈，连 1 分都加不进去（说明所有科目都满分了），强制退出
                // 避免死循环
                if (!adjusted) break;
            }
        }

        private static double NextGaussian(Random rnd, double mean, double stddev)
        {
            double u1 = 1.0 - rnd.NextDouble();
            double u2 = 1.0 - rnd.NextDouble();
            double z0 = Math.Sqrt(-2.0 * Math.Log(u1)) * Math.Cos(2 * Math.PI * u2);
            return mean + z0 * stddev;
        }
    }
}