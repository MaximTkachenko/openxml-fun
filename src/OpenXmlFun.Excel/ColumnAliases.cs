namespace OpenXmlFun.Excel
{
    internal static class ColumnAliases
    {
        private const string Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        public static readonly string[] ExcelColumnNames = new string[Alphabet.Length * 2];

        static ColumnAliases()
        {
            for (int i = 0; i < Alphabet.Length; i++)
            {
                ExcelColumnNames[i] = Alphabet[i].ToString();
            }

            for (int i = 0; i < Alphabet.Length; i++)
            {
                ExcelColumnNames[i + Alphabet.Length] = $"{Alphabet[0]}{Alphabet[i]}";
            }
        }
    }
}
