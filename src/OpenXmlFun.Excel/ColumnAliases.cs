namespace OpenXmlFun.Excel
{
    internal static class ColumnAliases
    {
        private const string Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        public static readonly string[] ExcelColumnNames = new string[Alphabet.Length + Alphabet.Length * Alphabet.Length];

        static ColumnAliases()
        {
            for (int i = 0; i < Alphabet.Length; i++)
            {
                ExcelColumnNames[i] = Alphabet[i].ToString();
            }

            for (int i = 0; i < Alphabet.Length; i++)
            for (int j = 0; j < Alphabet.Length; j++)
            {
                ExcelColumnNames[Alphabet.Length + Alphabet.Length * i + j] = $"{Alphabet[i]}{Alphabet[j]}";
            }
        }
    }
}
