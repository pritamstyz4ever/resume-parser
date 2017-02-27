# resume-parser
A resume parser built with Microsoft's OpenXML SDK

# Dependencies
Microsoft Open XML SDK: nuget package link https://www.nuget.org/packages/DocumentFormat.OpenXml/
Newton soft json serialization library: https://www.nuget.org/packages/newtonsoft.json/



# Parser
Uses WordprocessingDocument class to parse .docx files
Reads headers, linksa, and content inside resume
Parser is and independent library with uses a configurable keyword dictionary for searching content
Parser returns a json object with key values for keys configured in dictionary xml and corresponding content.

# Viewer
A ASP.NET MVC5 project using razor view to upload resume and display after parsing (pretty basic UI, nothing too fancy)

Please feel free to fork and make the code more robust and bug free.

# Note
Parser cannot be 100% accurate as there is no standard template for resume other than use of certain keywords.
Hence dictionary is maintained to make it configurable and extensible. 
It can further be exptended to identify specifc skillset from a list of documents.
