using System;
using System.Threading;

namespace MSOfficeManager.API
{
    public class MSOfficeDocument : IDisposable
    {
        internal dynamic App;
        internal dynamic Doc;

        /// <summary>
        /// Path file
        /// - Путь к файлу
        /// </summary>
        public string Path { get; private set; }

        /// <summary>
        /// Read Only status
        /// - Статус открытия для чтения
        /// </summary>
        public bool IsReadOnly { get; private set; }

        /// <summary>
        /// Block document close status
        /// - Статус блокировки закрытия документа
        /// </summary>
        public bool IsBlock { get; private set; }


        /// <summary>
        /// Document
        /// - Документ
        /// </summary>
        /// <param name="path"></param>
        /// <param name="app"></param>
        /// <param name="doc"></param>
        /// <param name="isReadOnly"></param>
        internal MSOfficeDocument(string path, object app, object doc, bool isReadOnly)
        {
            IsBlock = true;
            App = app;
            Doc = doc;
            Path = path;
            IsReadOnly = isReadOnly;
        }

        /// <summary>
        /// Close
        /// - Закрыть
        /// </summary>
        public void Close(bool save)
        {
            if (!IsReadOnly && save) Save();
            IsBlock = false;
            Doc?.Close(false);
            //Marshal.FinalReleaseComObject(Doc);
            App?.CloseDoc(Path);
            App = null;
            Doc = null;
            Path = null;
        }

        /// <summary>
        /// Close
        /// - Закрыть
        /// </summary>
        public void Dispose()
        {
            Close(false);
        }

        /// <summary>
        /// Save
        /// - Сохранить
        /// </summary>
        public void Save()
        {
            if (!IsReadOnly)
                Doc?.Save();
        }
    }
}
