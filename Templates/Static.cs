using System;
using System.Linq;

namespace MSOfficeManager.Templates
{
    internal static class Static
    {
        public static void FormatRow(Cell[] cells, bool renum)
        {
            cells = cells.OrderBy(x => x.Column).ToArray();
            for (int i = 0; i < cells.Length; i++)
                if (cells[i].Column > -1 && cells[i].LastMergeColumn > cells[i].Column)
                {
                    Cell c = cells[i];
                    int mc = cells[i].LastMergeColumn;
                    int ci = mc - c.Column;
                    for (int j = i + 1; j < cells.Length; j++)
                    {
                        if (cells[j].Column <= mc)
                        {
                            int lmc = cells[j].LastMergeColumn;
                            if (lmc > mc)
                            {
                                mc = lmc;
                                c.ColumnSpan = mc - c.Column + 1;
                            }
                            cells[j].Column = -1;
                        }
                        else
                            break;
                    }
                }
            if (renum)
                for (int i = 0; i < cells.Length; i++)
                    if (cells[i].Column > -1 && cells[i].LastMergeColumn > cells[i].Column)
                    {
                        int mc = cells[i].LastMergeColumn;
                        int l = mc - cells[i].Column;
                        for (int j = i + 1; j < cells.Length; j++)
                            if (cells[j].Column > mc)
                                cells[j].Column -= l;
                    }
        }

        public static char[] GetControls()
        {
            return new char[] {' ', '\0', '\u0001', '\u0002', '\u0003', '\u0004', '\u0005', '\u0006', '\a', '\b', '\t', '\n', '\v', '\f', '\r',
                '\u000e', '\u000f', '\u0010', '\u0011', '\u0012', '\u0013', '\u0014', '\u0015', '\u0016', '\u0017', '\u0018', '\u0019', '\u001a', '\u001b', '\u001c', '\u001d', '\u001e',
                '\u001f', '\u007f', '\u0080', '\u0081', '\u0082', '\u0083', '\u0084', '\u0085', '\u0086', '\u0087', '\u0088', '\u0089', '\u008a', '\u008b', '\u008c', '\u008d', '\u008e',
                '\u008f', '\u0090', '\u0091', '\u0092', '\u0093', '\u0094', '\u0095', '\u0096', '\u0097', '\u0098', '\u0099', '\u009a', '\u009b', '\u009c', '\u009d', '\u009e', '\u009f' };
        }

        /// <summary>
        /// Получить число из строки игнорируя символы кроме цифр
        /// </summary>
        /// <param name="input"></param>
        /// <param name="IsNegative"></param>
        /// <returns></returns>
        public static double GetDoubleInString(this string input, bool IsNegative)
        {
            if (String.IsNullOrEmpty(input))
                return 0;
            string ns = "";
            int si = 0;
            if (IsNegative && input[0] == '-')
            {
                ns = "-";
                si++;
            }
            for (int i = si; i < input.Length; i++)
                if (Char.IsDigit(input[i]))
                    ns += input[i];
                else if ((input[i] == ',' || input[i] == '.') && !ns.Contains(','))
                    ns += ',';

            double res = 0;
            if (Double.TryParse(ns, out res))
                return res;
            ns = ns.Replace(',', '.');
            if (Double.TryParse(ns, out res))
                return res;
            return 0;
        }
    }
}
