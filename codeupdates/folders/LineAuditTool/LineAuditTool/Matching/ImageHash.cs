// ImageHash.cs
// 64-bit dHash and pHash over 256x256 binarized image.
// C# 7.3.

using System;
using System.Drawing;

namespace LineAuditTool.Matching
{
    internal static class ImageHash
    {
        public static HashFeatures ComputeHashes(Bitmap binarized256)
        {
            HashFeatures f;
            f.DHash64 = DHash64(binarized256);
            f.PHash64 = PHash64(binarized256);
            return f;
        }

        // 64-bit dHash: compare adjacent pixels in a downscaled 9x8 grayscale
        public static ulong DHash64(Bitmap img)
        {
            const int W = 9, H = 8;
            using (Bitmap small = new Bitmap(W, H))
            using (Graphics g = Graphics.FromImage(small))
            {
                g.DrawImage(img, new Rectangle(0, 0, W, H));
                ulong hash = 0UL;
                int bit = 0;
                int y, x;
                for (y = 0; y < H; y++)
                {
                    for (x = 0; x < W - 1; x++)
                    {
                        int l1 = small.GetPixel(x, y).R;     // already 0/255
                        int l2 = small.GetPixel(x + 1, y).R;
                        if (l1 > l2) hash |= (1UL << bit);
                        bit++;
                    }
                }
                return hash;
            }
        }

        // 64-bit pHash using 32x32 DCT, take top-left 8x8 (excluding DC) vs mean
        public static ulong PHash64(Bitmap img)
        {
            const int N = 32;
            double[,] a = new double[N, N];
            using (Bitmap small = new Bitmap(N, N))
            using (Graphics g = Graphics.FromImage(small))
            {
                g.DrawImage(img, new Rectangle(0, 0, N, N));
                int y, x;
                for (y = 0; y < N; y++)
                {
                    for (x = 0; x < N; x++)
                    {
                        a[y, x] = small.GetPixel(x, y).R / 255.0; // 0..1
                    }
                }
            }

            double[,] dct = Dct2D(a);

            // Compute mean of 8x8 block excluding (0,0)
            int i, j;
            double sum = 0.0; int cnt = 0;
            for (i = 0; i < 8; i++)
            {
                for (j = 0; j < 8; j++)
                {
                    if (i == 0 && j == 0) continue;
                    sum += dct[i, j];
                    cnt++;
                }
            }
            double mean = (cnt > 0) ? sum / cnt : 0.0;

            ulong hash = 0UL;
            int bit = 0;
            for (i = 0; i < 8; i++)
            {
                for (j = 0; j < 8; j++)
                {
                    if (i == 0 && j == 0) continue;
                    if (dct[i, j] > mean) hash |= (1UL << bit);
                    bit++;
                }
            }
            return hash;
        }

        public static int Hamming64(ulong a, ulong b)
        {
            ulong x = a ^ b;
            int c = 0;
            while (x != 0) { x &= (x - 1); c++; }
            return c;
        }

        private static double[,] Dct2D(double[,] a)
        {
            int N = a.GetLength(0);
            int M = a.GetLength(1);
            double[,] outp = new double[N, M];

            double[] cN = new double[N];
            int i, j, u, v;
            for (i = 0; i < N; i++) cN[i] = (i == 0) ? 1.0 / Math.Sqrt(2.0) : 1.0;

            for (u = 0; u < N; u++)
            {
                for (v = 0; v < M; v++)
                {
                    double sum = 0.0;
                    for (i = 0; i < N; i++)
                    {
                        for (j = 0; j < M; j++)
                        {
                            sum += a[i, j] *
                                   Math.Cos(((2.0 * i + 1.0) * u * Math.PI) / (2.0 * N)) *
                                   Math.Cos(((2.0 * j + 1.0) * v * Math.PI) / (2.0 * M));
                        }
                    }
                    outp[u, v] = 0.25 * cN[u] * cN[v] * sum;
                }
            }
            return outp;
        }
    }
}
