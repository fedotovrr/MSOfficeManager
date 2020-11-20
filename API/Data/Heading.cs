using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace MSOfficeManager.API
{
    /// <summary>
    /// Заголовок структуры документа
    /// </summary>
    [Serializable]
    public class Heading
    {
        /// <summary>
        /// Наименование
        /// </summary>
        [XmlAttribute]
        public string Caption;

        /// <summary>
        /// Количество страниц или номер страницы
        /// </summary>
        [XmlAttribute]
        public int PageOrPages;

        public Heading()
        {

        }

        public Heading(string caption, int pageOrPages)
        {
            Caption = caption;
            PageOrPages = pageOrPages;
        }
    }
}
