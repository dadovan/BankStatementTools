using System.Drawing;

namespace BankStatementTools
{
    /// <summary>
    /// Stores text content pulled from the PDF along with position information
    /// </summary>
    public class TextObject
    {
        /// <summary>
        /// The position of the text relative to the PDF page
        /// </summary>
        public readonly PointF Position;

        /// <summary>
        /// The text content
        /// </summary>
        public readonly string Text;

        /// <summary>
        /// Constructs a new TextObject instance
        /// </summary>
        /// <param name="position">The position of the text relative to the PDF page</param>
        /// <param name="text">The text content</param>
        public TextObject(PointF position, string text)
        {
            Position = position;
            Text = text;
        }
    }
}