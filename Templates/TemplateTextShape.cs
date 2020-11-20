using System;
using System.Linq;

namespace MSOfficeManager.Templates
{
    /// <summary>
    /// Шаблон заполнения текстового блока (надписи)
    /// </summary>
    public class TemplateTextShape
    {
        /// <summary>
        /// Template value - Значение шаблона
        /// </summary>
        public string Value;

        /// <summary>
        /// Function identify by shape name - Функция идентификаци по имени
        /// </summary>
        public Func<string, bool> IdentifyByName { get; private set; }

        /// <summary>
        /// Function identify by shape text content - Функция идентификаци по тексту
        /// </summary>
        public Func<string, bool> IdentifyByContent { get; private set; }


        /// <summary>
        /// Шаблон заполнения текстового блока (надписи)
        /// </summary>
        /// <param name="value">значение шаблона</param>
        /// <param name="identify">функция идентификации</param>
        public TemplateTextShape(string value, Func<string, bool> identifyByContent, Func<string, bool> identifyByName = null)
        {
            Value = value;
            IdentifyByContent = identifyByContent;
            IdentifyByName = identifyByName;
        }

        /// <summary>
        /// Значение шаблона
        /// </summary>
        /// <returns></returns>
        public virtual string GetValue()
        {
            return Value;
        }
    }
}
