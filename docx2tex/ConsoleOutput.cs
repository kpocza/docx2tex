using System;
using System.Collections.Generic;
using System.Text;
using docx2tex.Library;

namespace docx2tex
{
    class ConsoleOutput : IStatusInformation
    {
        public void Write(string data)
        {
            Console.Write(data);
        }

        public void WriteCR(string data)
        {
            Console.Write(data + "\r");
        }

        public void WriteLine(string data)
        {
            Console.WriteLine(data);
        }

        public void WriteLine()
        {
            Console.WriteLine();
        }
    }
}
