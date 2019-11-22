using System;
using System.Text;

namespace HelperLibrary
{
    /// <summary>
    /// Implements useful string operations.
    /// </summary>
    public static class SupportStrings
    {
        /// <summary>
        /// Returns string with no consecutive CRLF pairs.
        /// </summary>
        /// <param name="value">Source string.</param>
        /// <returns>String without consecutive CRLF pairs.</returns>
        public static string RemoveDuplicatingLineBreaks(string value)
        {
            value = value.Replace("\r", string.Empty);
            string[] parts = value.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
            value = string.Join("\n", parts);
            return value.Replace("\n", "\r\n");
        }

        /// <summary>
        /// Joins an array of strings into single string using specified separator.
        /// </summary>
        /// <param name="separator">Separator string.</param>
        /// <param name="values">Array of strings to join.</param>
        /// <returns>New combined string.</returns>
        public static string Join(string separator, string[] values)
        {
            StringBuilder builder = new StringBuilder();

            for (int i = 0; i < values.Length; i++)
            {
                if (!string.IsNullOrWhiteSpace(values[i]))
                {
                    builder.Append(separator);
                    builder.Append(values[i].Trim());
                }
            }

            if (builder.Length > separator.Length)
            {
                builder.Remove(0, separator.Length);
            }

            return builder.ToString();
        }

        /// <summary>
        /// Replaces straight quotes ("Example") with paired quotes («Example»).
        /// </summary>
        /// <param name="text">String that contains straight quotes.</param>
        /// <returns>String that contains paired quotes.</returns>
        public static string ReplaceQuotes(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            char[] chars = text.ToCharArray();

            System.Collections.Generic.List<int> indexes = new System.Collections.Generic.List<int>();

            for (int i = 0; i < chars.Length; i++)
            {
                if (chars[i] == '"')
                {
                    indexes.Add(i);
                }
            }

            int border = indexes.Count % 2 == 0 ? indexes.Count / 2 : (indexes.Count / 2) + 1;

            for (int i = 0; i < indexes.Count; i++)
            {
                if (i < border)
                {
                    chars[indexes[i]] = '«';
                }
                else
                {
                    chars[indexes[i]] = '»';
                }
            }
           

            return new string(chars);
        }

        /// <summary>
        /// Corrects text formatting. Ensures correct spacing between words and punctuation marks (including decimal separator in numbers).
        /// </summary>
        /// <param name="text">Original text.</param>
        /// <returns>Corrected text.</returns>
        public static string FormatText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            StringBuilder result = new StringBuilder();

            char whitespace = ' ';

            text = '~' + text + '~';

            bool space = false, ignoreSpaces = true;
            char c, nextChar, prevChar;
            int j;
            for (int i = 1; i < text.Length - 1; i++)
            {
                c = text[i];

                if (c == whitespace)
                {
                    space = !ignoreSpaces;

                    continue;
                }

                if (CharIsOpeningBracket(c))
                {
                    result.Append(whitespace);
                    result.Append(c);

                    space = false;
                    ignoreSpaces = true;

                    continue;
                }

                if (CharIsClosingBracket(c))
                {
                    result.Append(c);

                    space = true;

                    continue;
                }

                ignoreSpaces = false;

                if (CharIsPunctuationMark(c))
                {
                    // Next non-whitespace character
                    j = i + 1;
                    do
                    {
                        nextChar = text[j];
                        j++;
                    }
                    while (nextChar == whitespace);

                    // Previous non-whitespace character
                    j = i - 1;
                    do
                    {
                        prevChar = text[j];
                        j--;
                    }
                    while (prevChar == whitespace);

                    if ((c == ',') && (prevChar == ','))
                    {
                        // Several commas in a row
                    }
                    else
                    {
                        result.Append(c);
                    }

                    if (((c == '.') || (c == ',')) && CharIsDigit(prevChar) && CharIsDigit(nextChar))
                    {
                        // This is a decimal separator in a number    
                        space = false;
                        ignoreSpaces = true;
                    }
                    else
                    {
                        space = true;
                    }

                    continue;
                }

                if (space)
                {
                    result.Append(whitespace);
                }

                result.Append(c);

                space = false;
            }

            return result.ToString();
        }

        /// <summary>
        /// Improves phone number formatting (for Russia only). Ensures correct spacing between parts.
        /// </summary>
        /// <param name="phoneNumber">Original phone number.</param>
        /// <returns>Improved phone number.</returns>
        public static string FormatPhoneNumber(string phoneNumber)
        {
            if (string.IsNullOrEmpty(phoneNumber))
            {
                return phoneNumber;
            }

            string result = phoneNumber.Trim();

            if (result.StartsWith("+"))
            {
                result = result.Remove(0, 1);
            }

            string[] words = result.Split(new char[] { '(', ')' }, StringSplitOptions.None);
            if (words.Length != 3)
            {
                // Phone number does not contain city code
                return phoneNumber;
            }

            #region Country code

            string countryCodeText = string.IsNullOrEmpty(words[0]) ? "7" : words[0].Trim();
            int countryCode;
            if (!int.TryParse(countryCodeText, out countryCode))
            {
                // Country code is not a number
                return phoneNumber;
            }

            if (countryCode == 8)
            {
                // Replace internal code with international code (in Russia)
                countryCode = 7;
            }

            #endregion

            #region City code

            string cityCodeText = string.IsNullOrEmpty(words[1]) ? null : words[1].Trim();
            if (cityCodeText == null)
            {
                // No city code
                return phoneNumber;
            }

            int cityCode;
            if (!int.TryParse(cityCodeText, out cityCode))
            {
                // City code is not a number
                return phoneNumber;
            }

            #endregion

            #region Phone number

            string numberText = string.IsNullOrEmpty(words[2]) ? null : words[2].Trim().Replace("-", string.Empty).Replace(" ", string.Empty);
            if (numberText == null)
            {
                return phoneNumber;
            }

            int number;
            if (!int.TryParse(numberText, out number))
            {
                // Phone number is not a number
                return phoneNumber;
            }

            numberText = number.ToString();

            switch (numberText.Length)
            {
                case 7:
                    numberText = numberText.Substring(0, 3) + "-" + numberText.Substring(3, 2) + "-" + numberText.Substring(5, 2);
                    break;
                case 6:
                    numberText = numberText.Substring(0, 3) + "-" + numberText.Substring(3, 3);
                    break;
                case 5:
                    numberText = numberText.Substring(0, 3) + "-" + numberText.Substring(3, 2);
                    break;
            }

            #endregion

            StringBuilder builder = new StringBuilder();
            builder.Append("+");
            builder.Append(countryCode);
            builder.Append(" (");
            builder.Append(cityCode);
            builder.Append(") ");
            builder.Append(numberText);

            return builder.ToString();
        }

        /// <summary>
        /// Prepares text for searching. Converts letters "ё" to "е".
        /// </summary>
        /// <param name="text">Original text.</param>
        /// <returns>Searchable text.</returns>
        public static string PrepareTextForSearching(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }
            else
            {
                return text.Trim().ToLower().Replace('ё', 'е');
            }
        }

        /// <summary>
        /// Returns string with repeated character.
        /// </summary>
        /// <param name="c">Character to repeat.</param>
        /// <param name="count">Number of repeats.</param>
        /// <returns>String with repeated character.</returns>
        public static string RepeatChar(char c, int count)
        {
            return new string(c, count);
        }

        private static bool CharIsPunctuationMark(char value)
        {
            switch (value)
            {
                case '.':
                case ',':
                case ':':
                case ';':
                case '?':
                case '!':
                    return true;
                default:
                    return false;
            }
        }

        private static bool CharIsOpeningBracket(char value)
        {
            switch (value)
            {
                case '(':
                case '[':
                case '{':
                    return true;
                default:
                    return false;
            }
        }

        private static bool CharIsClosingBracket(char value)
        {
            switch (value)
            {
                case ')':
                case ']':
                case '}':
                    return true;
                default:
                    return false;
            }
        }

        private static bool CharIsDigit(char value)
        {
            switch (value)
            {
                case '0':
                case '1':
                case '2':
                case '3':
                case '4':
                case '5':
                case '6':
                case '7':
                case '8':
                case '9':
                    return true;
                default:
                    return false;
            }
        }
    }
}
