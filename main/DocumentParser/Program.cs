using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentParser.src.parser;


namespace DocumentParser
{
    /// <summary>
    /// 
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// Defines the entry point of the application.
        /// </summary>
        /// <param name="args">The arguments.</param>
        private static void Main(string[] args)
        {

            const string documentPath = @"C:\Laptop Backup\Pritam\Docs\Resume_Pritam Paul.docx";
            var dictionaryFilePath = Path.Combine(Directory.GetCurrentDirectory(), @"resources\KeywordDictionary.xml");

            var response = Engine.Parse(documentPath, dictionaryFilePath);

            Console.WriteLine(response);
        }

    }

}
