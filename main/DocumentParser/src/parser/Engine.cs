using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentParser.src.models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace DocumentParser.src.parser
{
    public static class Engine
    {
        /// <summary>
        /// Parses the specified path.
        /// </summary>
        /// <param name="documentPath">The document path.</param>
        /// <param name="dictionaryPath">The dictionary path.</param>
        /// <returns></returns>
        public static string Parse(string documentPath, string dictionaryPath)
        {
            var wordDocument = XDocument.Load(dictionaryPath);

            var xElement = wordDocument.Element("Dictionary");
            if (xElement == null)
            {
                throw new Exception("Dictionary element not specified");
            }

            var wordDictionary = xElement.Elements()
                                         .ToDictionary(a => a.Name.ToString(), b => b.Elements().Nodes().ToList());

            using (var document = WordprocessingDocument.Open(documentPath, false))
            {
                ProcessHeaders(document);
                var body = ProcessBody(document, wordDictionary);
                var hyperlinks = ProcessHyperlinks(document, wordDictionary);

                var result = body.Concat(hyperlinks).GroupBy(d => d.Key).ToDictionary(d => d.Key, d => d.First().Value);

                return JsonConvert.SerializeObject(result);
            }

        }

        /// <summary>
        /// Parses the and serialize.
        /// </summary>
        /// <param name="documentPath">The document path.</param>
        /// <param name="dictionaryPath">The dictionary path.</param>
        /// <returns></returns>
        public static string ParseAndSerialize(string documentPath, string dictionaryPath)
        {
            var wordDocument = XDocument.Load(dictionaryPath);

            var xElement = wordDocument.Element("Dictionary");
            if (xElement == null)
            {
                throw new Exception("Dictionary element not specified");
            }

            var wordDictionary = xElement.Elements()
                                         .ToDictionary(a => a.Name.ToString(), b => b.Elements().Nodes().ToList());

            using (var document = WordprocessingDocument.Open(documentPath, false))
            {
                ProcessHeaders(document);
                var body = ProcessBody(document, wordDictionary);
                var jBody = JObject.Parse(JsonConvert.SerializeObject(body));
                var hyperlinks = ProcessHyperlinks(document, wordDictionary);
                var jLinks = JObject.Parse(JsonConvert.SerializeObject(hyperlinks));
                jBody.Merge(jLinks);

                return JsonConvert.SerializeObject(jBody);
            }
        }


        /// <summary>
        /// Processes the hyper-links.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <param name="wordDictionary">The word dictionary.</param>
        private static Dictionary<string, IList<string>> ProcessHyperlinks(WordprocessingDocument document, Dictionary<string, List<XNode>> wordDictionary)
        {
            var mainDocument = document.MainDocumentPart.Document;
            var links = new Dictionary<string, IList<string>>();

            if (!mainDocument.Descendants<Hyperlink>().Any()) return links;

            foreach (var hyperlink in
                mainDocument.Descendants<Hyperlink>())
            {
                var hyperlinkText = new StringBuilder();

                foreach (var text in hyperlink.Descendants<Text>())
                {
                    hyperlinkText.Append(text.InnerText);
                }

                var link = hyperlinkText.ToString();

                string key;

                ContainsAnyContent(link, wordDictionary, out key);

                if (links.ContainsKey("Handle"))
                {
                    var value = links["Handle"];
                    value.Add(link);
                }
                else
                {
                    links.Add(key, new List<string> {link});
                }

            }

            return links;
        }

        /// <summary>
        /// Processes the headers.
        /// </summary>
        /// <param name="document">The document.</param>
        private static void ProcessHeaders(WordprocessingDocument document)
        {
            var mainDoc = document.MainDocumentPart.Document;
            var headers = mainDoc.Descendants<HeaderReference>().Count();
            //TODO: parse headers if required
        }

        /// <summary>
        /// Processes the body.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <param name="wordDictionary">The word dictionary.</param>
        private static Dictionary<string, IList<string>> ProcessBody(WordprocessingDocument document, Dictionary<string, List<XNode>> wordDictionary)
        {
            var body = document.MainDocumentPart.Document.Body;

            var defaultStyleName =
                document.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>()
                        .First(style => style.Type == StyleValues.Paragraph &&
                                        style.Default)
                        .StyleId.Value;


            var styles = body.Elements().Select((p, i) =>
                {
                    var styleNode = p.Descendants<ParagraphStyleId>().FirstOrDefault();
                    var styleName = styleNode != null ? styleNode.Val.Value : defaultStyleName;

                    return new WordDocumentModel
                        {
                            Element = p,
                            Index = i,
                            StyleName = styleName
                        };
                });

            var documentModel = styles.Select(i =>
                {
                    string text;

                    if (i.Element.GetType() == typeof(Paragraph))
                        text = i.Element
                                .Descendants<Text>()
                                .Where(z => z.Parent is Run || z.Parent is InsertedRun)
                                .StringConcatenate(element => element.Text);
                    else
                        text = i.Element
                                .Descendants<Paragraph>()
                                .StringConcatenate(p => p.Descendants<Text>()
                                                         .Where(z => z.Parent is Run || z.Parent is InsertedRun)
                                                         .StringConcatenate(element => element.Text),
                                                   Environment.NewLine);

                    return new WordDocumentModel
                        {
                            StyleName = i.StyleName,
                            Index = i.Index,
                            Text = text
                        };
                }).ToList();

            var key = string.Empty;
            var headers =
                documentModel.Where(i => ContainsAnyContent(i.Text, wordDictionary, out key) && i.StyleName.Contains("Heading"))
                             .Select(i => new HeaderModel
                                 {
                                     HeaderName = key,
                                     HeaderIndex = i.Index
                                 }).ToList();

            var summary = new Dictionary<string, IList<string>>();

            ProcessElements(summary, documentModel, headers.GetEnumerator());

            return summary;
        }

        /// <summary>
        /// Processes the elements.
        /// </summary>
        /// <param name="summary">The summary.</param>
        /// <param name="documentModel">The document model.</param>
        /// <param name="iterator">The iterator.</param>
        private static void ProcessElements(IDictionary<string, IList<string>> summary,
                                            List<WordDocumentModel> documentModel, List<HeaderModel>.Enumerator iterator)
        {
            var current = iterator.Current;

            if (current != null)
            {
                if (summary.ContainsKey(current.HeaderName))
                {
                    return;
                }
                summary.Add(current.HeaderName, null);

                if (iterator.MoveNext())
                {
                    var previous = current;
                    current = iterator.Current;

                    if (current != null)
                        summary[previous.HeaderName] = GetContent(documentModel, (previous.HeaderIndex + 1),
                                                                  (current.HeaderIndex - (previous.HeaderIndex + 1)));
                    ProcessElements(summary, documentModel, iterator);
                }
                else
                {
                    var lastIndex = documentModel.Last().Index;
                    summary[current.HeaderName] = GetContent(documentModel, current.HeaderIndex + 1,
                                                             (lastIndex - (current.HeaderIndex + 1)));

                }

            }
            if (iterator.MoveNext())
            {
                ProcessElements(summary, documentModel, iterator);
            }
        }

        /// <summary>
        /// Gets the content.
        /// </summary>
        /// <param name="documentModel">The document model.</param>
        /// <param name="previous">The previous.</param>
        /// <param name="current">The current.</param>
        /// <returns></returns>
        private static List<string> GetContent(IEnumerable<WordDocumentModel> documentModel, int previous, int current)
        {
            return documentModel.Skip(previous)
                                .Take(current)
                                .Select(a => a.Text)
                                .ToList();
        }

        /// <summary>
        /// Determines whether [contains any content] [the specified string to search].
        /// </summary>
        /// <param name="stringToSearch">The string to search.</param>
        /// <param name="nodes">The search strings.</param>
        /// <param name="key">The key.</param>
        /// <returns></returns>
        private static bool ContainsAnyContent(string stringToSearch, Dictionary<string, List<XNode>> nodes, out string key)
        {
            key = string.Empty;

            foreach (var node in nodes.Where(node => node.Value.Any(n => stringToSearch.ToLower().Contains(n.ToString()))))
            {
                key = node.Key;
                return true;
            }

            return false;
        }

        /// <summary>
        /// Strings the concatenate.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source">The source.</param>
        /// <param name="func">The function.</param>
        /// <returns></returns>
        private static string StringConcatenate<T>(this IEnumerable<T> source, Func<T, string> func)
        {
            var sb = new StringBuilder();
            foreach (var item in source)
                sb.Append(func(item));
            return sb.ToString();
        }

        /// <summary>
        /// Strings the concatenate.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source">The source.</param>
        /// <param name="func">The function.</param>
        /// <param name="separator">The separator.</param>
        /// <returns></returns>
        private static string StringConcatenate<T>(this IEnumerable<T> source,
                                                  Func<T, string> func, string separator)
        {
            var sb = new StringBuilder();
            foreach (var item in source)
                sb.Append(func(item)).Append(separator);
            if (sb.Length > separator.Length)
                sb.Length -= separator.Length;
            return sb.ToString();
        }

    }
}
