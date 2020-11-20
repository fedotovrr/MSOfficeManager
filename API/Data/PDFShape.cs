using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace MSOfficeManager.API
{
    [Serializable]
    public class PDFShape
    {
        /// <summary>
        /// Имя
        /// </summary>
        [XmlAttribute]
        public string Name;

        /// <summary>
        /// Файл
        /// </summary>
        [XmlAttribute]
        public byte[] File;

        public PDFShape(string name, byte[] file)
        {
            Name = name;
            File = file;
        }
    }
}
