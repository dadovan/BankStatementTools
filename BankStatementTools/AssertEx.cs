using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace BankStatementTools
{
    /// <summary>
    /// Extensions to the Assert class for use in unit tests
    /// </summary>
    public static class AssertEx
    {
        /// <summary>
        /// Allows an action that generates an exception to be performed.  This method wraps the action and validates the exception type and message.
        /// <paramref name="message"/> is optional but may include a '*' wildcard at the start or end of the text.  It is case-insensitive.
        /// Ex:
        ///   "*should not be null" will match "The parameter xyz should not be null"
        ///   "Expected 'test' but received*" will match "Expected 'test' but recived 'some data'"
        ///   "*was not executed*" will match "The operation was not executed properly"
        /// </summary>
        /// <typeparam name="T">The type of <see cref="Exception"/> to expect</typeparam>
        /// <param name="action">The action expected to generate the exception</param>
        /// <param name="message">Optionally, text used to validate the exception's <see cref="Exception.Message"/> property</param>
        public static void Throws<T>(Action action, string message = null) where T : Exception
        {
            Assert.IsNotNull(action, $"{nameof(action)} can't be null");
            try
            {
                action();
            }
            catch (Exception exception)
            {
                if (typeof(T) != exception.GetType())
                    throw new AssertFailedException($"An exception of type {typeof(T).Name} was expected but received: {exception.GetType().Name}, {exception.Message}");
                if ((!String.IsNullOrWhiteSpace(message)) && (!TestExceptionMessage(message, exception.Message)))
                    throw new AssertFailedException($"An exception of type {typeof(T).Name} was caught.  Expected message [{message}] but received [{exception.Message}].");
                // TODO: For unit tests, Console.WriteLine is fine.  If we want to use this in product code, we need to support more flexible logging.
                Console.WriteLine($"Expected exception caught: {typeof(T).Name}, {exception.Message}");
                return;
            }

            throw new AssertFailedException($"An exception of type {typeof(T).Name} was not thrown");
        }

        // TODO: Could make into generic wildcard method
        /// <summary>
        /// Tests to see if exception messages match a wildcard-able string
        /// <paramref name="message"/> may include a '*' wildcard at the start or end of the text.  It is case-insensitive.
        /// Ex:
        ///   "*should not be null" will match "The parameter xyz should not be null"
        ///   "Expected 'test' but received*" will match "Expected 'test' but recived 'some data'"
        ///   "*was not executed*" will match "The operation was not executed properly"
        /// </summary>
        /// <param name="message">The test string</param>
        /// <param name="actual">The actual exception message</param>
        /// <returns>True if the exception message matches <paramref name="message"/></returns>
        private static bool TestExceptionMessage(string message, string actual)
        {
            Assert.IsFalse(String.IsNullOrWhiteSpace(message), $"{nameof(message)} can't be null");
            Assert.IsFalse(String.IsNullOrWhiteSpace(actual), $"{nameof(actual)} can't be null");
            var wildStart = message[0] == '*';
            var wildEnd = message[message.Length - 1] == '*';
            if (wildStart && wildEnd)
            {
                var test = message.Substring(1, message.Length - 2);
                return actual.IndexOf(test, 0, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            if (wildStart)
            {
                var test = message.Substring(1);
                return actual.EndsWith(test, StringComparison.OrdinalIgnoreCase);
            }
            if (wildEnd)
            {
                var test = message.Substring(0, message.Length - 1);
                return actual.StartsWith(test, StringComparison.OrdinalIgnoreCase);
            }
            return actual.Equals(message, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Compares a pair of float arrays, providing more debug info than CollectionAssert.AreEqual() does.
        /// </summary>
        /// <param name="expected">The expected array</param>
        /// <param name="actual">The actual array</param>
        /// <param name="logAllData">True to log all the elements from both arrays</param>
        public static void CompareCollection(float[] expected, float[] actual, bool logAllData = false)
        {
            Assert.IsNotNull(expected, $"{nameof(expected)} can't be null");
            Assert.IsNotNull(actual, $"{nameof(actual)} can't be null");
            if (logAllData)
            {
                // TODO: For unit tests, Console.WriteLine is fine.  If we want to use this in product code, we need to support more flexible logging.
                Console.WriteLine("Expected:");
                foreach (var s in expected.OrderBy(v => v))
                    Console.WriteLine(s);
                Console.WriteLine("Actual:");
                foreach (var s in actual.OrderBy(v => v))
                    Console.WriteLine(s);
            }
            foreach (var s in expected)
                Assert.IsTrue(actual.Contains(s), $"Unable to find {s} in actual");
            Assert.AreEqual(expected.Length, actual.Length);
        }
    }

    namespace Test
    {
        [TestClass]
        public class AssertExTests
        {
            [TestMethod]
            public void ThrowsTest()
            {
                AssertEx.Throws<ArgumentException>(() => throw new ArgumentException());
                AssertEx.Throws<ArgumentException>(() => throw new ArgumentException("exact match"), "exact match");
                AssertEx.Throws<ArgumentException>(() => throw new ArgumentException("CaSe InSeNsItIvE"), "case insensitive");
                AssertEx.Throws<ArgumentException>(() => throw new ArgumentException("test end with"), "*end with");
                AssertEx.Throws<ArgumentException>(() => throw new ArgumentException("start with test"), "start with*");
                AssertEx.Throws<ArgumentException>(() => throw new ArgumentException("somewhere in the middle i think"), "*in the middle*");
            }

            [TestMethod]
            public void ThrowsActionNullTest()
            {
                AssertEx.Throws<AssertFailedException>(() => AssertEx.Throws<Exception>(null), "*action can't be null");
            }

            [TestMethod]
            public void ThrowsNegativeTest()
            {
                AssertEx.Throws<AssertFailedException>(() => AssertEx.Throws<AssertFailedException>(() => { }), "*an exception of type AssertFailedException was not thrown");
                AssertEx.Throws<AssertFailedException>(() => AssertEx.Throws<Exception>(() => throw new Exception("hello"), "*goodbye*"),
                    "*Expected message [*goodbye*] but received [hello].");
                AssertEx.Throws<AssertFailedException>(() => AssertEx.Throws<AssertFailedException>(() => throw new Exception("hello"), "*goodbye*"),
                    "*An exception of type AssertFailedException was expected but received: Exception, hello");
            }
        }
    }
}
