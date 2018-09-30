using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PdfSharp.Pdf.IO;

namespace BankStatementTools
{
    /// <summary>
    /// A reader for Synovus statements
    /// </summary>
    public class SynovusStatementReader
    {
        // TODO: English only for now...
        private const string AmountText = "Amount";
        private const string BalanceSummaryText = "Balance Summary";
        private const string ChecksText = "Checks";
        private const string DepositsOtherCreditsText = "Deposits/Other Credits";
        private const string LastStatementText = "Last statement:";
        private const string OtherDebitsText = "Other Debits";
        private const string ThisStatementText = "This statement:";
        private const string TransactionTypeText = "Transaction Type";

        private const float DefaultComparisonPrecision = 0.001f;

        private List<Transaction> m_transactions;

        /// <summary>
        /// This statement's date
        /// </summary>
        public DateTime StatementDate { get; set; }

        /// <summary>
        /// A read-only list of <see cref="Transaction"/>s in this statement
        /// </summary>
        public ReadOnlyCollection<Transaction> Transactions => new ReadOnlyCollection<Transaction>(m_transactions);

        /// <summary>
        /// The total amount of credits in this statement
        /// </summary>
        public float Credits => m_transactions.Where(t => t.Amount > 0f).Sum(t => t.Amount);
        /// <summary>
        /// The total amount of debits in this statement
        /// </summary>
        public float Debits => m_transactions.Where(t => t.Amount < 0f).Sum(t => t.Amount);

        /// <summary>
        /// Private constructor to enforce construction via <see cref="Load"/> method.
        /// </summary>
        private SynovusStatementReader() { }

        /// <summary>
        /// Constructs a new <see cref="SynovusStatementReader"/> instance based on <paramref name="filename"/>.
        /// </summary>
        /// <param name="filename">The statement to load</param>
        /// <returns>A <see cref="SynovusStatementReader"/> instance representing the file</returns>
        public static SynovusStatementReader Load(string filename)
        {
            Assert.IsNotNull(filename, $"{nameof(filename)} can't be null.");
            Assert.IsTrue(File.Exists(filename), $"File does not exist: {filename}");
            var pdfdoc = PdfReader.Open(filename, PdfDocumentOpenMode.Import);
            Assert.AreEqual("Synovus Bank", pdfdoc.Info.Author);
            var pages = Enumerable.Range(0, pdfdoc.PageCount).Select(i => pdfdoc.Pages[i].ExtractText().ToList()).ToArray();
            var statement = CreateStatement(pages);
            return statement;
        }

        /// <summary>
        /// Private constructor used by unit tests.
        /// </summary>
        /// <param name="pages">An array of <see cref="TextObject"/> <see cref="List{T}"/>s, one per statement page</param>
        /// <returns>A new <see cref="SynovusStatementReader"/> instance</returns>
        private static SynovusStatementReader Load(params List<TextObject>[] pages) => CreateStatement(pages);

        /// <summary>
        /// Constructs a new <see cref="SynovusStatementReader"/> instance based on a list of statement pages.
        /// </summary>
        /// <param name="pages">An array of <see cref="TextObject"/> <see cref="List{T}"/>s, one per statement page</param>
        /// <returns>A new <see cref="SynovusStatementReader"/> instance</returns>
        private static SynovusStatementReader CreateStatement(params List<TextObject>[] pages)
        {
            Assert.IsNotNull(pages, $"{nameof(pages)} can't be null");
            Assert.IsTrue(pages.Length > 0, $"{nameof(pages)} must contain at least one page");
            var beginningBalance = 0d;
            var credits = 0d;
            var debits = 0d;
            var endingBalance = 0d;
            var statement = new SynovusStatementReader();
            for (var i = 0; i < pages.Length; i++)
            {
                var textObjects = pages[i];
                Assert.IsNotNull(textObjects, $"Page {i} can't be null");
                if (i == 0)
                {
                    statement.StatementDate = FindStatementDate(textObjects);

                    // DEBUG:
                    // foreach (var to in textObjects)
                    //    Console.WriteLine($"({to.Position.X}, {to.Position.Y}) {to.Text}");

                    var checks = FindTextObject(textObjects, ChecksText);
                    var bb = FindTextObject(textObjects, "Beginning balance", onOrAboveY: checks.Position.Y);
                    var lb = FindTextObject(textObjects, "Low balance", onOrAboveY: checks.Position.Y);
                    var lbX = lb.Position.X;
                    var bal = FindTextObjects(textObjects, onOrAboveX: bb.Position.X + 0.01f, belowX: lbX, onOrAboveY: bb.Position.Y);
                    var balb = bal.Single(to => to.Position.Y == bb.Position.Y);
                    beginningBalance = Double.Parse(balb.Text);
                    Console.WriteLine(beginningBalance);

                    var dc = FindTextObject(textObjects, "Deposits/Credits", onOrAboveY: checks.Position.Y);
                    var v2 = FindTextObjects(textObjects, onOrAboveX: dc.Position.X + 0.01f, belowX: lbX, onOrAboveY: dc.Position.Y);
                    var v2t = v2.Single(to => to.Position.Y == dc.Position.Y);
                    credits = Double.Parse(v2t.Text);
                    Console.WriteLine(credits);

                    var wd = FindTextObject(textObjects, "Withdrawals/Debits", onOrAboveY: checks.Position.Y);
                    var v3 = FindTextObjects(textObjects, onOrAboveX: wd.Position.X + 0.01f, belowX: lbX, onOrAboveY: wd.Position.Y);
                    var v3t = v3.Single(to => to.Position.Y == wd.Position.Y);
                    debits = Double.Parse(v3t.Text);
                    Console.WriteLine(debits);

                    var eb = FindTextObject(textObjects, "Ending balance", onOrAboveY: checks.Position.Y, belowY: wd.Position.Y);
                    var v4 = FindTextObjects(textObjects, onOrAboveX: eb.Position.X + 0.01f, belowX: lbX, onOrAboveY: eb.Position.Y);
                    var v4t = v4.Single(to => to.Position.Y == eb.Position.Y);
                    endingBalance = Double.Parse(v4t.Text);
                    Console.WriteLine(endingBalance);

                    statement.m_transactions = new List<Transaction>();
                }
                var transactions = GetTransactions(textObjects, statement.StatementDate);
                if ((transactions != null) && (transactions.Count > 0))
                    statement.m_transactions.AddRange(transactions);
            }

            Assert.AreEqual(endingBalance, beginningBalance + credits - debits, DefaultComparisonPrecision, "Mismatch validating balance summary.  " +
                $"Beginning balance ({beginningBalance}) + credits ({credits}) - debits ({debits}) != ending balance {endingBalance}");
            Assert.AreEqual(credits, statement.Credits, DefaultComparisonPrecision);
            Assert.AreEqual(-debits, statement.Debits, DefaultComparisonPrecision);

            return statement;
        }

        /// <summary>
        /// Transforms one or more input PDF files into an output CSV file.
        /// </summary>
        /// <param name="outputFilename">The name of the output file</param>
        /// <param name="inputFilenames">The list of input files</param>
        public static void TransformStatements(string outputFilename, params string[] inputFilenames)
        {
            Assert.IsNotNull(outputFilename, $"{nameof(outputFilename)} can't be null");
            Assert.IsNotNull(inputFilenames, $"{nameof(inputFilenames)} can't be null");
            Assert.IsTrue(inputFilenames.Length > 0, $"{nameof(inputFilenames)} must contain at least one file");
            var allTransactions = new List<Transaction>();
            foreach (var inputFilename in inputFilenames)
            {
                Console.WriteLine($"Loading {inputFilename}");
                var statement = Load(inputFilename);

                var credits = statement.m_transactions.Where(t => t.Amount > 0f).Sum(t => t.Amount);
                var debits = statement.m_transactions.Where(t => t.Amount < 0f).Sum(t => t.Amount);

                Console.WriteLine($"\tStatement Date: {statement.StatementDate}");
                Console.WriteLine($"\tTransactions: {statement.m_transactions.Count}");
                Console.WriteLine($"\tCredits: {credits}");
                Console.WriteLine($"\tDebits: {debits}");

                allTransactions.AddRange(statement.m_transactions);
            }

            var csvLines = allTransactions.BuildCsvRows();
            File.WriteAllLines(outputFilename, csvLines, Encoding.Unicode);
        }

        /// <summary>
        /// Finds a single <see cref="TextObject"/> matching the given characteristics.
        /// </summary>
        /// <remarks>
        /// Note that PDF pages have their origin at the bottom left corner!
        /// </remarks>
        /// <param name="text">If not null, the text to match (case in-sensitive)</param>
        /// <param name="onOrAboveX">If > 0, the X of the <see cref="TextObject"/> must be above this.</param>
        /// <param name="belowX">If > 0, the X of the <see cref="TextObject"/> must be below this.</param>
        /// <param name="onOrAboveY">If > 0, the Y of the <see cref="TextObject"/> must be above this.</param>
        /// <param name="belowY">If > 0, the Y of the <see cref="TextObject"/> must be below this.</param>
        /// <param name="throwIfEmpty">True to throw if no <see cref="TextObject"/>s were found.</param>
        private static TextObject FindTextObject(IEnumerable<TextObject> textObjects, string text = null, float onOrAboveX = -1f, float belowX = -1f,
            float onOrAboveY = -1f, float belowY = -1f, bool throwIfEmpty = true)
        {
            Assert.IsNotNull(textObjects, $"{nameof(textObjects)} must not be null");
            var textObject = FindTextObjects(textObjects, text, onOrAboveX, belowX, onOrAboveY, belowY, throwIfEmpty).SingleOrDefault();
            if (throwIfEmpty && (textObject == null))
                ThrowFindTextObjectException(text, onOrAboveX, belowX, onOrAboveY, belowY);
            return textObject;
        }

        /// <summary>
        /// Generates a descriptive exception message and throws a <see cref="KeyNotFoundException"/> for errors coming from FindTextObject* calls.
        /// </summary>
        /// <param name="text">If not null, the text to match (case in-sensitive)</param>
        /// <param name="onOrAboveX">If > 0, the X of the <see cref="TextObject"/> must be above this.</param>
        /// <param name="belowX">If > 0, the X of the <see cref="TextObject"/> must be below this.</param>
        /// <param name="onOrAboveY">If > 0, the Y of the <see cref="TextObject"/> must be above this.</param>
        /// <param name="belowY">If > 0, the Y of the <see cref="TextObject"/> must be below this.</param>
        private static void ThrowFindTextObjectException(string text, float onOrAboveX, float belowX, float onOrAboveY, float belowY)
        {
            var message = String.Format("text: {0}, onOrAboveX: {1}, belowX: {2}, onOrAboveY: {3}, belowY: {4}",
                String.IsNullOrWhiteSpace(text) ? "(none)" : text,
                onOrAboveX < 0f ? "(none)" : onOrAboveX.ToString(CultureInfo.InvariantCulture),
                belowX < 0f ? "(none)" : belowX.ToString(CultureInfo.InvariantCulture),
                onOrAboveY < 0f ? "(none)" : onOrAboveY.ToString(CultureInfo.InvariantCulture),
                belowY < 0f ? "(none)" : belowY.ToString(CultureInfo.InvariantCulture));
            throw new KeyNotFoundException($"Unable to find TextObject matching: '{message}'");
        }

        /// <summary>
        /// Finds all <see cref="TextObject"/>s matching the given characteristics.
        /// </summary>
        /// <remarks>
        /// Note that PDF pages have their origin at the bottom left corner!
        /// </remarks>
        /// <param name="textObjects">The list of <see cref="TextObject"/>s to search</param>
        /// <param name="text">If not null, the text to match (case in-sensitive)</param>
        /// <param name="onOrAboveX">If > 0, the X of the <see cref="TextObject"/> must be above this.</param>
        /// <param name="belowX">If > 0, the X of the <see cref="TextObject"/> must be below this.</param>
        /// <param name="onOrAboveY">If > 0, the Y of the <see cref="TextObject"/> must be above this.</param>
        /// <param name="belowY">If > 0, the Y of the <see cref="TextObject"/> must be below this.</param>
        /// <param name="throwIfEmpty">True to throw if no <see cref="TextObject"/>s were found.</param>
        /// <returns>Any matching <see cref="TextObject"/>s</returns>
        private static List<TextObject> FindTextObjects(IEnumerable<TextObject> textObjects, string text = null, float onOrAboveX = -1f, float belowX = -1f,
            float onOrAboveY = -1f, float belowY = -1f, bool throwIfEmpty = true)
        {
            Assert.IsNotNull(textObjects, $"{nameof(textObjects)} must not be null");
            var matchingTextObjects = textObjects
                .Where(to => (onOrAboveX < 0) || to.Position.X >= onOrAboveX)
                .Where(to => (belowX < 0) || to.Position.X < belowX)
                .Where(to => (onOrAboveY < 0) || to.Position.Y >= onOrAboveY)
                .Where(to => (belowY < 0) || to.Position.Y < belowY)
                .Where(to => String.IsNullOrWhiteSpace(text) || to.Text.Equals(text, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (throwIfEmpty && (matchingTextObjects.Count == 0))
                ThrowFindTextObjectException(text, onOrAboveX, belowX, onOrAboveY, belowY);
            return matchingTextObjects;
        }

        /// <summary>
        /// Finds the statement date.
        /// </summary>
        /// <param name="textObjects">The list of <see cref="TextObject"/>s to search</param>
        /// <returns>The statement date</returns>
        private static DateTime FindStatementDate(List<TextObject> textObjects)
        {
            Assert.IsNotNull(textObjects, $"{nameof(textObjects)} must not be null");
            var thisStatementObject = FindTextObject(textObjects, ThisStatementText);
            var lastStatementObject = FindTextObject(textObjects, LastStatementText);
            var statementDateObject = FindTextObject(textObjects, onOrAboveX: thisStatementObject.Position.X + 0.01f, belowX: 500f,
                onOrAboveY: thisStatementObject.Position.Y, belowY: lastStatementObject.Position.Y);
            return DateTime.Parse(statementDateObject.Text);
        }

        /// <summary>
        /// Creates a list of all <see cref="Transaction"/>s found in the <see cref="TextObject"/>s.
        /// </summary>
        /// <param name="textObjects">The list of <see cref="TextObject"/>s to search</param>
        /// <param name="statementDate">The statement date</param>
        /// <returns>All <see cref="Transaction"/>s</returns>
        private static List<Transaction> GetTransactions(List<TextObject> textObjects, DateTime statementDate)
        {
            Assert.IsNotNull(textObjects, $"{nameof(textObjects)} must not be null");

            var transactions = new List<Transaction>();

            (float bottomY, float topY) = FindDebitsSectionBounds(textObjects);
            if ((bottomY >= 0f) && (topY >= 0f))
                transactions.AddRange(GetTransactions(textObjects, statementDate, bottomY, topY, areCredits: false));

            (bottomY, topY) = FindCreditsSectionBounds(textObjects);
            if ((bottomY >= 0f) && (topY >= 0f))
                transactions.AddRange(GetTransactions(textObjects, statementDate, bottomY, topY, areCredits: true));

            // TODO: This logic fails if checks spills over onto the next page.  Not fixing yet and the final credit/debit total checks catch this.
            var checks = FindTextObject(textObjects, ChecksText, throwIfEmpty: false);
            var otherDebits = FindTextObject(textObjects, OtherDebitsText, throwIfEmpty: false);
            if ((checks != null) && (otherDebits != null))
            {
                bottomY = otherDebits.Position.Y + 0.01f;
                var amount = FindTextObjects(textObjects, AmountText, onOrAboveY: bottomY, belowY: checks.Position.Y).First();
                topY = amount.Position.Y;
                transactions.AddRange(GetChecks(textObjects, statementDate, bottomY, topY));
            }

            return transactions;
        }

        private static List<Transaction> GetTransactions(List<TextObject> textObjects, DateTime statementDate, float bottomY, float topY, bool areCredits)
        {
            var allTransactionObjects = textObjects.Where(to => (to.Position.Y > bottomY) && (to.Position.Y < topY)).ToList();

            var dates = allTransactionObjects.Where(to => Regex.IsMatch(to.Text, @"^\d\d\-\d\d$")).ToList();
            var rowBounds = dates.Select(to => to.Position.Y).OrderByDescending(y => y).ToList();
            rowBounds.Add(bottomY);

            var transactions = new List<Transaction>();
            for (var i = 0; i < rowBounds.Count - 1; i++)
            {
                var top = rowBounds[i];
                var bottom = rowBounds[i + 1];
                var transactionObjects = allTransactionObjects.Where(to => (to.Position.Y > bottom) && (to.Position.Y <= top)).ToList();
                var xs = transactionObjects.Select(to => to.Position.X).OrderBy(x => x).Distinct().ToList();

                var dateObject = transactionObjects.Single(to => to.Position.X == xs[0]);
                var date = DateTime.Parse($"{dateObject.Text}-{statementDate.Year}");

                var transactionTypeObject = transactionObjects.Single(to => to.Position.X == xs[1]);
                var transactionType = transactionTypeObject.Text;

                string[] description;
                if (xs.Count == 4)
                {
                    var descriptionObjects = transactionObjects.Where(to => to.Position.X == xs[2]).OrderByDescending(to => to.Position.Y);
                    description = descriptionObjects.Select(to => to.Text).ToArray();
                }
                else
                    description = null; // Some transactions have no description, like certain deposits or service charges

                var amountObject = transactionObjects.Single(to => to.Position.X > xs[xs.Count - 2]);
                var amount = Single.Parse(amountObject.Text);
                if (!areCredits)
                    amount = -amount;

                transactions.Add(new Transaction(date, transactionType, description, amount));
            }
            return transactions;
        }

        private static List<Transaction> GetChecks(List<TextObject> textObjects, DateTime statementDate, float bottomY, float topY)
        {
            var allTransactionObjects = FindTextObjects(textObjects, onOrAboveY: bottomY, belowY: topY).Where(to => !to.Text.StartsWith("*")).ToList();
            var dates = allTransactionObjects.Where(to => Regex.IsMatch(to.Text, @"^\d\d\-\d\d$")).ToList();
            var rows = dates.Select(d => d.Position.Y).Distinct().ToArray();

            var transactions = new List<Transaction>();
            foreach (var row in rows)
            {
                var allElements = allTransactionObjects.Where(to => to.Position.Y == row).OrderBy(to => to.Position.X).ToArray();
                var leftElements = allElements.Take(3).ToArray();
                var rightElements = allElements.Length > 3 ? allElements.Skip(3).ToArray() : null;

                foreach (var elements in new[] { leftElements, rightElements })
                {
                    if (elements == null)
                        continue;
                    var date = DateTime.Parse($"{elements[1].Text}-{statementDate.Year}");
                    var description = new[] { "Check #" + elements[0].Text.TrimEnd('*', ' ') };
                    var amount = -(Single.Parse(elements[2].Text));

                    transactions.Add(new Transaction(date, "Check", description, amount));
                }
            }
            return transactions;
        }

        /// <summary>
        /// Finds the horizontal boundary for the Other Debits transaction sections
        /// </summary>
        /// <param name="textObjects">The full set of TextObjects to consider.  (Note these cannot span pages)</param>
        /// <returns>The bottom and top Y positions</returns>
        private static (float bottomY, float topY) FindDebitsSectionBounds(IReadOnlyCollection<TextObject> textObjects)
        {
            Assert.IsNotNull(textObjects, $"{nameof(textObjects)} can't be null");

            var otherDebits = FindTextObject(textObjects, OtherDebitsText, throwIfEmpty: false);
            if (otherDebits == null)
                return (-1f, -1f);
            var depositsOtherCredits = FindTextObject(textObjects, DepositsOtherCreditsText, belowY: otherDebits.Position.Y, throwIfEmpty: false);
            var bottomY = depositsOtherCredits?.Position.Y ?? 0f;
            if (bottomY <= 0f)
            {
                var balanceSummary = FindTextObject(textObjects, BalanceSummaryText, belowY: otherDebits.Position.Y, throwIfEmpty: false);
                bottomY = balanceSummary?.Position.Y ?? 0f;
            }

            var transactionType = FindTextObject(textObjects, TransactionTypeText, onOrAboveY: bottomY, belowY: otherDebits.Position.Y);

            var topY = transactionType.Position.Y;
            return (bottomY, topY);
        }

        /// <summary>
        /// Finds the horizontal boundary for the Other Debits transaction sections
        /// </summary>
        /// <param name="textObjects">The full set of TextObjects to consider.  (Note these cannot span pages)</param>
        /// <returns>The bottom and top Y positions</returns>
        private static (float bottomY, float topY) FindCreditsSectionBounds(IReadOnlyCollection<TextObject> textObjects)
        {
            Assert.IsNotNull(textObjects, $"{nameof(textObjects)} can't be null");

            var depositsOtherCredits = FindTextObject(textObjects, DepositsOtherCreditsText, throwIfEmpty: false);
            if (depositsOtherCredits == null)
                return (-1f, -1f);

            var balanceSummary = FindTextObject(textObjects, BalanceSummaryText, belowY: depositsOtherCredits.Position.Y, throwIfEmpty: false);
            var bottomY = balanceSummary?.Position.Y ?? 0f;

            var transactionType = FindTextObject(textObjects, TransactionTypeText, onOrAboveY: bottomY, belowY: depositsOtherCredits.Position.Y);

            var topY = transactionType.Position.Y;
            return (bottomY, topY);
        }

        /// <summary>
        /// Logs all <see cref="TextObject"/>s to the Console.  Helper method, used to debug.
        /// </summary>
        /// <param name="textObjects">The set of <see cref="TextObject"/>s to log</param>
        private static void LogTextObjects(IEnumerable<TextObject> textObjects)
        {
            foreach (var textObject in textObjects)
                Console.WriteLine($"({textObject.Position.X}, {textObject.Position.Y}) {textObject.Text}");
        }
    }

    namespace Test
    {
        [TestClass]
        public class SynovusStatementTests
        {
            private const float DefaultComparisonPrecision = 0.001f;

            private List<TextObject> m_sampleTextObjects;

            private static List<TextObject> GetSampleTextObjects()
            {
                const int expectedTextObjects = 109;
                var beginningBalance = 2000f;
                var credits = 0f;
                var debits = 1298.82;
                var endingBalance = 701.18;
                var data = $@"(83.01, 647.5) JANE DOE
(335.01, 680.9) Page 1 of 1
(953.01, 720) PAGE 1
(1053.01, 720) 48317603814
(1153.01, 720) 0004
(1253.01, 720) 340
(1353.01, 720) 001007
(1453.01, 720) Y
(1553.01, 720) N
(1653.01, 720) 0
(83.01, 636.4) 1234 EXAMPLE ROAD
(83.01, 625.3) SOMEWHERE ND 87126-0000
(221.01, 514.4) Summary of Account Balance
(59.01, 425.7) Pro Business Checking
(335.01, 747.2) Statement of Account
(331.67, 292.6) * Skip in check sequence
(59.01, 492.2) Account
(293.01, 492.2) Number
(465.71, 492.2) Ending Balance
(97.96, 337) Number          Date
(246.31, 337) Amount
(59.01, 248.3) Date
(107.01, 248.3) Transaction Type
(233.01, 248.3) Description
(510.31, 248.3) Amount
(335.01, 725.1) Last statement:
(425.01, 725.1) June 30, 2016
(335.01, 714) This statement:
(425.01, 714) July 31, 2016
(335.01, 702.9) Total days in statement period: 31
(335.01, 691.8) 965-392-224-7
(437.01, 691.8) 624
(473.01, 691.8) 851
(331.96, 337) Number          Date
(480.31, 337) Amount
(59.01, 470) Pro Business Checking
(281.01, 470) 965-392-224-7
(483.23, 470) $0.00
(281.01, 425.7) Account Number  965-392-224-7
(474.74, 425.7) 6 Enclosures
(59.65, 403.5) Beginning balance
(253, 403.5) {beginningBalance:F2}
(59.65, 392.4) Deposits/Credits
(243.34, 392.4) 0.00
(323.01, 392.4) Low balance
(499, 392.4) 0.00
(59.65, 381.3) Withdrawals/Debits
(243.34, 381.3) {debits:F2}
(323.01, 381.3) Average balance
(499, 381.3) 0.00
(59.65, 370.3) Ending balance
(243.43, 370.3) {endingBalance:F2}
(323.01, 370.3) Average collected balance
(499, 370.3) 0.05
(53.01, 226.1) 07-02
(101.01, 226.1) Check Card Purchase
(233.01, 226.1) Merchant Purchase Terminal 819579
(517.2, 226.1) 17.36
(53.01, 192.8) 07-03
(101.01, 192.8) Check Card Purchase
(233.01, 192.8) Merchant Purchase Terminal 819578
(517.2, 192.8) 21.19
(53.01, 159.6) 07-05
(101.01, 159.6) Preauthorized Wd
(233.01, 159.6) Paypal Inst Xfer
(517.2, 159.6) 48.20
(53.01, 137.4) 07-06
(101.01, 137.4) Preauthorized Wd
(233.01, 137.4) Paypal Inst Xfer
(510.98, 137.4) 160.13
(53.01, 115.2) 07-07
(101.01, 115.2) Check Card Purchase
(233.01, 115.2) Merchant Purchase Terminal 517241
(517.2, 115.2) 19.77
(53.01, 82) 07-07
(101.01, 82) Check Card Purchase
(233.01, 82) Merchant Purchase Terminal 819578
(517.2, 82) 22.16
(98.33, 314.8) 1049
(161.01, 314.8) 07-05
(246.98, 314.8) 108.00
(98.33, 303.7) 1051 *
(161.01, 303.7) 07-01
(246.98, 303.7) 313.00
(98.33, 292.6) 1052
(161.01, 292.6) 07-13
(246.98, 292.6) 108.00
(98.33, 281.6) 1055 *
(161.01, 281.6) 07-19
(246.98, 281.6) 313.00
(332.33, 314.8) 1056
(395.01, 314.8) 07-26
(480.98, 314.8) 108.00
(332.33, 303.7) 1059 *
(395.01, 303.7) 07-26
(487.2, 303.7) 60.01
(335.01, 647.5) Direct inquiries to:
(335.01, 636.4) 888 123-5555
(53.01, 348.1) Checks
(53.01, 259.4) Other Debits
(233.01, 215) COURTYARD SARATOGA FL
(233.01, 203.9) TRAN DATE 07-08-17XXXXXXXXXXXX9112
(233.01, 181.8) COURTYARD SARATOGA FL
(233.01, 170.7) TRAN DATE 07-08-17XXXXXXXXXXXX9115
(233.01, 148.5) 827412
(233.01, 126.3) 827412
(233.01, 104.1) COURTYARD SARATOGA FL
(233.01, 93.1) TRAN DATE 07-08-17XXXXXXXXXXXX9213
(233.01, 70.9) COURTYARD SARATOGA FL
(233.01, 59.8) TRAN DATE 07-08-17XXXXXXXXXXXX9337";
                return ConvertToTextObjects(data, expectedTextObjects);
            }

            private static List<TextObject> ConvertToTextObjects(string data, int expectedTextObjects)
            {
                var textObjects = new List<TextObject>();
                foreach (Match m in Regex.Matches(data, @"^\((?<x>([\d\.]+)),\s(?<y>([\d\.]+))\)\s(?<text>(.+))\r$", RegexOptions.Multiline))
                {
                    var x = Single.Parse(m.Groups["x"].Value);
                    var y = Single.Parse(m.Groups["y"].Value);
                    var text = m.Groups["text"].Value;
                    textObjects.Add(new TextObject(new PointF(x, y), text));
                }
                Assert.AreEqual(expectedTextObjects, textObjects.Count);
                Assert.IsNotNull(textObjects.SingleOrDefault(to => to.Text.Equals("Other Debits", StringComparison.OrdinalIgnoreCase)));
                return textObjects;
            }

            // ********** Access members for private members of class
            private string LastStatementText => new PrivateType(typeof(SynovusStatementReader)).GetStaticField("LastStatementText") as string;
            private string ThisStatementText => new PrivateType(typeof(SynovusStatementReader)).GetStaticField("ThisStatementText") as string;

            private static SynovusStatementReader Load(params List<TextObject>[] pages)
            {
                var pt = new PrivateType(typeof(SynovusStatementReader));
                var statement = pt.InvokeStatic(nameof(Load), new Object[] { pages }) as SynovusStatementReader;
                return statement;
            }

            private static TextObject FindTextObject(IEnumerable<TextObject> textObjects, string text = null, float onOrAboveX = -1f, float belowX = -1f,
                float onOrAboveY = -1f, float belowY = -1f, bool throwIfEmpty = true)
            {
                var pt = new PrivateType(typeof(SynovusStatementReader));
                var textObject = pt.InvokeStatic(nameof(FindTextObject), textObjects, text, onOrAboveX, belowX, onOrAboveY, belowY, throwIfEmpty) as TextObject;
                return textObject;
            }

            private static List<TextObject> FindTextObjects(IEnumerable<TextObject> textObjects, string text = null, float onOrAboveX = -1f, float belowX = -1f,
                float onOrAboveY = -1f, float belowY = -1f, bool throwIfEmpty = true)
            {
                var pt = new PrivateType(typeof(SynovusStatementReader));
                var objects = pt.InvokeStatic(nameof(FindTextObjects), textObjects, text, onOrAboveX, belowX, onOrAboveY, belowY, throwIfEmpty) as List<TextObject>;
                return objects;
            }

            private static DateTime FindStatementDate(List<TextObject> textObjects)
            {
                var pt = new PrivateType(typeof(SynovusStatementReader));
                var date = (DateTime)pt.InvokeStatic(nameof(FindStatementDate), textObjects);
                return date;
            }

            private static (float bottomY, float topY) FindDebitsSectionBounds(IReadOnlyCollection<TextObject> textObjects)
            {
                var pt = new PrivateType(typeof(SynovusStatementReader));
                var tuple = (ValueTuple<float, float>)pt.InvokeStatic(nameof(FindDebitsSectionBounds), textObjects);
                return (tuple.Item1, tuple.Item2);
            }

            // ********** Test methods
            [TestInitialize]
            public void TestInitialize()
            {
                m_sampleTextObjects = GetSampleTextObjects();
            }

            [TestMethod]
            public void FindTextObjectTest()
            {
                var textObject = FindTextObject(m_sampleTextObjects, "Merchant Purchase Terminal 819579");
                Assert.AreEqual(233.01f, textObject.Position.X, DefaultComparisonPrecision);
                Assert.AreEqual(226.1f, textObject.Position.Y, DefaultComparisonPrecision);

                textObject = FindTextObject(m_sampleTextObjects, "Merchant Purchase Terminal 819579", onOrAboveY: 226.1f);
                Assert.AreEqual(233.01f, textObject.Position.X, DefaultComparisonPrecision);
                Assert.AreEqual(226.1f, textObject.Position.Y, DefaultComparisonPrecision);

                AssertEx.Throws<KeyNotFoundException>(() => FindTextObject(m_sampleTextObjects, "Merchant Purchase Terminal 123456", onOrAboveY: 249f),
                    "Unable to find TextObject matching*");
            }


            [TestMethod]
            public void FindTextObjectsTest()
            {
                var textObject = FindTextObject(m_sampleTextObjects, ThisStatementText);
                Assert.AreEqual(335.01f, textObject.Position.X, DefaultComparisonPrecision);
                Assert.AreEqual(714f, textObject.Position.Y, DefaultComparisonPrecision);

                textObject = FindTextObject(m_sampleTextObjects, LastStatementText);
                Assert.AreEqual(335.01f, textObject.Position.X, DefaultComparisonPrecision);
                Assert.AreEqual(725.1f, textObject.Position.Y, DefaultComparisonPrecision);

                var textObjects = FindTextObjects(m_sampleTextObjects, onOrAboveY: 714f);
                Assert.AreEqual(13, textObjects.Count);

                textObjects = FindTextObjects(m_sampleTextObjects, onOrAboveY: 714f, onOrAboveX: 1000f);
                Assert.AreEqual(7, textObjects.Count);

                textObjects = FindTextObjects(m_sampleTextObjects, onOrAboveX: 335.02f, belowX: 500f, onOrAboveY: 714f, belowY: 725.1f);
                Assert.AreEqual(1, textObjects.Count);
                Assert.AreEqual(425.01f, textObjects[0].Position.X, DefaultComparisonPrecision);
                Assert.AreEqual(714f, textObjects[0].Position.Y, DefaultComparisonPrecision);
            }

            [TestMethod]
            public void FindTransactionSectionBoundsTest()
            {
                (float bottomY, float topY) = FindDebitsSectionBounds(m_sampleTextObjects);
                Assert.AreEqual(248.3, topY, DefaultComparisonPrecision);
                Assert.AreEqual(0, bottomY, DefaultComparisonPrecision);
            }

            [TestMethod]
            public void FindStatementDateTest()
            {
                var date = FindStatementDate(m_sampleTextObjects);
                Assert.AreEqual("7/31/2016", date.ToShortDateString());
            }

            [TestMethod]
            public void FindAllTransactionsTest()
            {
                var transactions = Load(m_sampleTextObjects).Transactions;
                Assert.AreEqual(12, transactions.Count);

                Assert.AreEqual("7/2/2016", transactions[0].Date.ToShortDateString());
                Assert.AreEqual("Check Card Purchase", transactions[0].TransactionType);
                Assert.AreEqual(3, transactions[0].Description.Length);
                Assert.AreEqual("Merchant Purchase Terminal 819579", transactions[0].Description[0]);
                Assert.AreEqual("COURTYARD SARATOGA FL", transactions[0].Description[1]);
                Assert.AreEqual("TRAN DATE 07-08-17XXXXXXXXXXXX9112", transactions[0].Description[2]);
                Assert.AreEqual(-17.36f, transactions[0].Amount);

                Assert.AreEqual("7/7/2016", transactions[5].Date.ToShortDateString());
                Assert.AreEqual("Check Card Purchase", transactions[5].TransactionType);
                Assert.AreEqual(2, transactions[5].Description.Length);
                Assert.AreEqual("Merchant Purchase Terminal 819578", transactions[5].Description[0]);
                Assert.AreEqual("COURTYARD SARATOGA FL", transactions[5].Description[1]);
                Assert.AreEqual(-22.16f, transactions[5].Amount);

                Assert.AreEqual("7/19/2016", transactions[11].Date.ToShortDateString());
                Assert.AreEqual("Check", transactions[11].TransactionType);
                Assert.AreEqual(1, transactions[11].Description.Length);
                Assert.AreEqual("Check #1055", transactions[11].Description[0]);
                Assert.AreEqual(-313f, transactions[11].Amount);
            }
        }
    }
}
