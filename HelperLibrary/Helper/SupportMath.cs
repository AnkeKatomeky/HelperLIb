using System;

namespace HelperLibrary
{
    /// <summary>
    /// Implements useful mathematic methods.
    /// </summary>
    public static class SupportMath
    {
        /// <summary>
        /// Returns value quantized by specified quantizer (i.e. largest value that is smaller than specified value and divisible by specified quantizer).
        /// </summary>
        /// <param name="value">Value to be quantized.</param>
        /// <param name="quantizer">Quantization value.</param>
        /// <returns>Quantized value.</returns>
        public static int Quantize(int value, int quantizer)
        {
            if (value * quantizer < 0)
            {
                throw new ArgumentException("Arguments have different signs");
            }

            if (quantizer == 0)
            {
                return value;
            }
            else
            {
                if (value % quantizer == 0)
                {
                    return value;
                }
                else
                {
                    return value / quantizer * quantizer;
                }
            }
        }        

        /// <summary>
        /// Округляет значение денежной суммы до целых копеек (2 знака после запятой).
        /// </summary>
        public static decimal RoundAsMoney(decimal value)
        {
            return decimal.Round(value, 2, MidpointRounding.AwayFromZero);
        }

        /// <summary>
        /// Возвращает сокращенное название единицы измерения (штука или упаковка).
        /// </summary>
        public static string GetUnitName(bool singleUnit)
        {
            return singleUnit ? "шт" : "упак";
        }
    }
}
