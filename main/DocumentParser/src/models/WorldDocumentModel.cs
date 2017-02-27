using DocumentFormat.OpenXml;
using Newtonsoft.Json;

namespace DocumentParser.src.models
{
       /// <summary>
    /// 
    /// </summary>
    public class WordDocumentModel
    {
        /// <summary>
        /// Gets or sets the element.
        /// </summary>
        /// <value>
        /// The element.
        /// </value>
        [JsonIgnore]
        public OpenXmlElement Element { get; set; }
        /// <summary>
        /// Gets or sets the name of the style.
        /// </summary>
        /// <value>
        /// The name of the style.
        /// </value>
        public string StyleName { get; set; }
        /// <summary>
        /// Gets or sets the index.
        /// </summary>
        /// <value>
        /// The index.
        /// </value>
        public int Index { get; set; }
        /// <summary>
        /// Gets or sets the text.
        /// </summary>
        /// <value>
        /// The text.
        /// </value>
        public string Text { get; set; }
    }
}
