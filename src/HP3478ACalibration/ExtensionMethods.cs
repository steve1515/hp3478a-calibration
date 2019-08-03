using System.Collections.Generic;
using System.Diagnostics;

namespace HP3478ACalibration
{
    public static class ExtensionMethods
    {
        [DebuggerStepThrough]
        public static IEnumerable<T> GetRow<T>(this T[,] array, int rowIndex)
        {
            for (int i = 0; i < array.GetLength(1); i++)
            {
                yield return array[rowIndex, i];
            }
        }
    }
}
