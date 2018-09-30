using System.Collections.Generic;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Content.Objects;

namespace BankStatementTools
{
    /// <summary>
    /// Helpers for the PdfSharp package
    /// </summary>
    public static class PdfSharpHelpers
    {
        /// <summary>
        /// Extracts all <see cref="TextObject"/>s from the PDF page
        /// </summary>
        /// <param name="page">The PDF page</param>
        /// <returns>All <see cref="TextObject"/>s</returns>
        public static IEnumerable<TextObject> ExtractText(this PdfPage page)
        {
            Assert.IsNotNull(page, $"{nameof(page)} can't be null");
            var content = ContentReader.ReadContent(page);
            var textObjects = content.ExtractText();
            return textObjects;
        }

        private static PointF m_lastPoint = PointF.Empty;
        // With help from https://stackoverflow.com/questions/10141143/c-sharp-extract-text-from-pdf-using-pdfsharp/23667589
        /// <summary>
        /// Extracts all <see cref="TextObject"/>s from the <see cref="CObject"/>, recursively if needed.
        /// </summary>
        /// <param name="cObject">The <see cref="CObject"/> to query</param>
        /// <returns>All <see cref="TextObject"/>s from the <see cref="CObject"/></returns>
        private static IEnumerable<TextObject> ExtractText(this CObject cObject)
        {
            Assert.IsNotNull(cObject, $"{nameof(cObject)} can't be null");
            if (cObject is COperator cOperator)
            {
                if (cOperator.OpCode.Name == OpCodeName.Td.ToString())
                {
                    Assert.AreEqual(2, cOperator.Operands.Count);
                    var x = (float)((cOperator.Operands[0] as CReal).Value);
                    var y = (float)((cOperator.Operands[1] as CReal).Value);
                    m_lastPoint = new PointF(x, y);
                }
                if (cOperator.OpCode.Name == OpCodeName.Tj.ToString() ||
                    cOperator.OpCode.Name == OpCodeName.TJ.ToString())
                {
                    foreach (var cOperand in cOperator.Operands)
                        foreach (var textObject in ExtractText(cOperand))
                            yield return textObject;
                }
            }
            else if (cObject is CSequence)
            {
                var cSequence = (CSequence)cObject;
                foreach (var element in cSequence)
                    foreach (var textObject in ExtractText(element))
                        yield return textObject;
            }
            else if (cObject is CString)
            {
                var cString = (CString)cObject;
                yield return new TextObject(m_lastPoint, cString.Value);
            }
        }
    }
}