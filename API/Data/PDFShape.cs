using System;
using System.Xml.Serialization;

namespace MSOfficeManager.API
{
    [Serializable]
    public class PDFShape
    {
        [XmlAttribute]
        public string Name { get; set; }

        [XmlAttribute]
        public byte[] File { get; set; }

        [XmlAttribute]
        public double dx { get; set; }

        [XmlAttribute]
        public double dy { get; set; }

        public PDFShape()
        {

        }

        public PDFShape(string name, byte[] file)
        {
            Name = name;
            File = file;
        }
    }
}
