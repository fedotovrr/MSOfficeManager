using System;
using System.Collections.Generic;
using System.Linq;

namespace MSOfficeManager.Templates
{
    /// <summary>
    /// Шаблон заполнения документа
    /// </summary>
    public class TemplateData
    {
        /// <summary>
        /// Таблицы
        /// </summary>
        public IEnumerable<TemplateTable> Tables;

        /// <summary>
        /// Текстовые блоки (надписи) документа
        /// </summary>
        public IEnumerable<TemplateTextShape> TextShapes;

        /// <summary>
        /// Текстовые блоки (надписи) верхних колонтитулов
        /// </summary>
        public IEnumerable<TemplateTextShape> HeaderTextShape;

        /// <summary>
        /// Текстовые блоки (надписи) нижних колонтитулов
        /// </summary>
        public IEnumerable<TemplateTextShape> FooterTextShapes;
    }
}
