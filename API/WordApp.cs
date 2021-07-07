using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;
using static MSOfficeManager.API.Static;
using System.Runtime.ExceptionServices;

namespace MSOfficeManager.API
{
    /// <summary>
    /// Word App
    /// - Приложение Word
    /// </summary>
    public class WordApp : IDisposable
    {
        private Application App;
        private List<Process> AppProcess = new List<Process>();
        private List<WordDoc> Docs = new List<WordDoc>();
        private string TryPath;
        private bool Visible;

        /// <summary>
        /// Word App
        /// - Приложение Word
        /// </summary>
        public WordApp(bool visible)
        {
            Init(visible);
        }


        /// <summary>
        /// Create Document
        /// - Создать документ
        /// </summary>
        /// <param name="path"></param>
        /// <param name="overwrite"></param>
        /// <returns></returns>
        [HandleProcessCorruptedStateExceptions]
        public WordDoc CreateDoc(string path, bool overwrite)
        {
            try
            {
                if (App == null) throw new Exception($"Не удалось создать документ {path} приложение не запущено");
                if (IsDocumentOpen(path)) throw new Exception($"Документ уже редактируется {path}");
                if (File.Exists(path))
                {
                    if (overwrite)
                        File.Delete(path);
                    else
                        throw new Exception($"Документ уже существует {path}");
                }
                TryPath = path;
                object missing = System.Reflection.Missing.Value;
                object opath = path;
                Document doc = App.Documents.Add();
                SwitchView(doc);
                if (Templates.Static.GetDoubleInString(App.Version, false) > 13) doc.SaveAs2(ref opath, WdSaveFormat.wdFormatStrictOpenXMLDocument, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                else doc.SaveAs(ref opath, WdSaveFormat.wdFormatDocument);
                WordDoc d = new WordDoc(path, this, doc, false);
                Docs.Add(d);
                TryPath = null;
                return d;
            }
            catch (Exception e)
            {
                Close();
                throw new Exception($"Не удалось создать документ {path} ", e);
            }
        }

        /// <summary>
        /// Open Document
        /// - Открыть документ
        /// </summary>
        /// <param name="path"></param>
        /// <param name="readOnly"></param>
        /// <returns></returns>
        [HandleProcessCorruptedStateExceptions]
        public WordDoc OpenDoc(string path, bool readOnly)
        {
            try
            {
                if (App == null) throw new Exception($"Не удалось открыть документ {path} приложение не запущено");
                if (!readOnly && IsDocumentOpen(path)) throw new Exception($"Документ уже редактируется {path}");
                TryPath = path;
                object opath = path;
                Document doc = App.Documents.Open(opath, Type.Missing, readOnly);
                if (!readOnly) SwitchView(doc);
                WordDoc d = new WordDoc(path, this, doc, readOnly);
                Docs.Add(d);
                TryPath = null;
                return d;
            }
            catch (Exception e)
            {
                Close();
                throw new Exception($"Не удалось открыть документ {path} ", e);
            }
        }

        /// <summary>
        /// Try Open Document
        /// - попытаться открыть документ
        /// </summary>
        /// <param name="path"></param>
        /// <param name="readOnly"></param>
        /// <returns></returns>
        [HandleProcessCorruptedStateExceptions]
        public WordDoc TryOpenDoc(string path, bool readOnly)
        {
            Document doc = null;
            try
            {
                if (App == null) throw new Exception($"Не удалось открыть документ {path} приложение не запущено");
                if (!readOnly && IsDocumentOpen(path)) throw new Exception($"Документ уже редактируется {path}");
                TryPath = path;
                object opath = path;
                doc = App.Documents.Open(opath, Type.Missing, readOnly);
                if (!readOnly) SwitchView(doc);
                WordDoc d = new WordDoc(path, this, doc, readOnly);
                Docs.Add(d);
                TryPath = null;
                return d;
            }
            catch (Exception e)
            {
                doc?.Close(false);
                TryPath = null;
                return null;
            }
        }

        /// <summary>
        /// Close Document
        /// - Закрыть документ
        /// </summary>
        /// <param name="path"></param>
        internal void CloseDoc(string path)
        {
            if (Docs.FirstOrDefault(x => x.Path == path) is WordDoc d)
            {
                if (d.IsBlock)
                    d.Close(false);
                else
                    Docs.Remove(d);
            }
        }

        /// <summary>
        /// Close App
        /// - Закрыть приложение
        /// </summary>
        [HandleProcessCorruptedStateExceptions]
        public void Close()
        {
            try
            {
                for (int i = 0; i < Docs.Count; i++)
                    Docs[i].Close(false);
                App.DocumentOpen -= App_DocumentOpen;
                App.DocumentBeforeClose -= App_DocumentBeforeClose;
                App.Quit();
                Marshal.FinalReleaseComObject(App);
                Thread.Sleep(300);
            }
            catch (Exception) { }
            Clear();
        }

        /// <summary>
        /// Close App
        /// - Закрыть приложение
        /// </summary>
        public void Dispose()
        {
            Close();
        }


        [HandleProcessCorruptedStateExceptions]
        private void Init(bool visible)
        {
            try
            {
                Visible = visible;
                if (!IsRegistred("WINWORD.EXE")) throw new Exception("Word не установлен");
                int[] ps1 = Process.GetProcessesByName("WinWord").Select(x => x.Id).ToArray();

                App = new Application();
                App.Visible = visible;

                int[] ps2 = Process.GetProcessesByName("WinWord").Select(x => x.Id).ToArray();
                for (int i = 0; i < ps2.Length; i++)
                    if (!ps1.Contains(ps2[i]))
                    {
                        Process p = Process.GetProcessById(ps2[i]);
                        if (p.MainWindowHandle.ToInt32() == 0)
                            AppProcess.Add(p);
                    }

                App.DocumentOpen += App_DocumentOpen;
                App.DocumentBeforeClose += App_DocumentBeforeClose;
            }
            catch (Exception e)
            {
                Close();
                throw e;
            }
        }

        private void SwitchView(Document doc)
        {
            foreach (Window item in doc.Windows)
            {
                if (item.View.SplitSpecial == WdSpecialPane.wdPaneNone)
                    item.ActivePane.View.Type = WdViewType.wdPrintView;
                else
                    item.View.Type = WdViewType.wdPrintView;
            }
        }

        private void Clear()
        {
            for (int i = 0; i < AppProcess.Count; i++)
            {
                try
                {
                    AppProcess[i].Kill();
                }
                catch (Exception) { }
            }
            App = null;
            AppProcess.Clear();
            Docs.Clear();
        }

        private bool IsDocumentOpen(string path)
        {
            try
            {
                List<ApplicationClass> Apps = FindApps.GetObjects(typeof(ApplicationClass)) as List<ApplicationClass>;
                if (Apps == null) return false;
                for (int i = 0; i < Apps.Count; i++)
                    for (int j = 1; j <= Apps[i].Documents.Count; j++)
                        if (Path.Combine(Apps[i].Documents[j].Path, Apps[i].Documents[j].Name) == path)
                            return true;
            }
            catch (Exception) { }
            return false;
        }

        private void App_DocumentBeforeClose(Document Doc, ref bool Cancel)
        {
            string dpath = Path.Combine(Doc.Path, Doc.Name);
            if (Docs.FirstOrDefault(x => x.IsBlock && x.Path == dpath) != null)
                Cancel = true;
        }

        [HandleProcessCorruptedStateExceptions]
        private void App_DocumentOpen(Document Doc)
        {
            try
            {
                string dpath = Path.Combine(Doc.Path, Doc.Name);
                bool readOnly = Doc.ReadOnly;
                App.Visible = Visible;
                if (TryPath != dpath)
                {
                    Doc.Close();
                    Application App2 = new Application();
                    object path2 = dpath;
                    App2.Documents.Open(path2, Type.Missing, readOnly);
                    App2.Visible = true;
                }
            }
            catch (Exception e)
            {
                Close();
                throw e;
            }
        }
    }
}
