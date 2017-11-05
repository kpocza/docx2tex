using System;
using System.Collections.Generic;
using System.Text;

namespace docx2tex.Library
{
    /// <summary>
    /// Interface for status info output
    /// </summary>
    public interface IStatusInformation
    {
        /// <summary>
        /// Write simple text
        /// </summary>
        /// <param name="data"></param>
        void Write(string data);

        /// <summary>
        /// Write simple text with CR
        /// </summary>
        /// <param name="data"></param>
        void WriteCR(string data);

        /// <summary>
        /// Write simple text with CRLF
        /// </summary>
        /// <param name="data"></param>
        void WriteLine(string data);
    }
}
