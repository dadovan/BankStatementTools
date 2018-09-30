using System;
using System.Collections.Generic;
using System.Linq;

namespace BankStatementTools
{
    public static class TransactionEx
    {
        /// <summary>
        /// Builds the set of CSV rows out of a set of <see cref="Transaction"/>s
        /// </summary>
        /// <param name="transactions">The <see cref="Transaction"/>s to use</param>
        /// <param name="addHeader">True to add a header line.  False to omit.</param>
        /// <returns>The set of CSV rows</returns>
        public static IEnumerable<string> BuildCsvRows(this IEnumerable<Transaction> transactions, bool addHeader = true)
        {
            var lines = transactions
                .OrderBy(t => t.Date)
                .Select(t =>
                {
                    var description = (t.Description == null ? t.TransactionType : String.Join("  ", t.Description)).Replace(',', ' ');
                    var line = String.Join(",", t.Date.ToShortDateString(), description, t.Amount);
                    return line;
                })
                .ToList();
            if (addHeader)
                lines.Insert(0, "Date,Description,Amount");
            return lines;
        }
    }
}
