using System;

namespace BankStatementTools
{
    /// <summary>
    /// Stores the information related to a single transaction
    /// </summary>
    public class Transaction
    {
        /// <summary>
        /// The date of the transaction.  Note: time information is not available.
        /// </summary>
        public readonly DateTime Date;

        /// <summary>
        /// The type of transaction
        /// </summary>
        public readonly string TransactionType;

        /// <summary>
        /// The transaction description.
        /// </summary>
        public readonly string[] Description;

        /// <summary>
        /// The transaction amount
        /// </summary>
        public readonly float Amount;

        /// <summary>
        /// Constructs a new Transaction instance
        /// </summary>
        /// <param name="date">The date of the transaction</param>
        /// <param name="transactionType">The type of transaction</param>
        /// <param name="description">The transaction description</param>
        /// <param name="amount">The transaction amount</param>
        public Transaction(DateTime date, string transactionType, string[] description, float amount)
        {
            Date = date;
            TransactionType = transactionType ?? throw new ArgumentNullException(nameof(transactionType));
            Description = description;
            Amount = amount;
        }
    }
}