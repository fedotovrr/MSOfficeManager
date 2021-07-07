using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;
using System.Runtime.ExceptionServices;
using static MSOfficeManager.API.Static;
using System.Reflection;

namespace MSOfficeManager.API
{
    /// <summary>
    /// Excel App
    /// - Приложение Excel
    /// </summary>
    public class ExcelApp : IDisposable
    {
        private Application App;
        private List<Process> AppProcess = new List<Process>();
        private List<ExcelDoc> Docs = new List<ExcelDoc>();
        private string TryPath;
        private bool Visible;

        public Application Application => App;


        /// <summary>
        /// Excel App
        /// - Приложение Excel
        /// </summary>
        public ExcelApp(bool visible)
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
        public ExcelDoc CreateDoc(string path, bool overwrite)
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
                object opath = path;
                Workbook doc = App.Workbooks.Add();
                if (Templates.Static.GetDoubleInString(App.Version, false) > 13) doc.SaveAs(opath, XlFileFormat.xlOpenXMLWorkbook, Missing.Value, Missing.Value, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlUserResolution, true, Missing.Value, Missing.Value, Missing.Value);
                else doc.SaveAs(opath, XlFileFormat.xlExcel8, Missing.Value, Missing.Value, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlUserResolution, true, Missing.Value, Missing.Value, Missing.Value);
                ExcelDoc d = new ExcelDoc(path, this, doc, false);
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
        public ExcelDoc OpenDoc(string path, bool readOnly)
        {
            try
            {
                if (App == null) throw new Exception($"Не удалось открыть документ {path} приложение не запущено");
                if (!readOnly && IsDocumentOpen(path)) throw new Exception($"Документ уже редактируется {path}");
                TryPath = path;
                Workbook doc = App.Workbooks.Open(path, Type.Missing, readOnly);
                ExcelDoc d = new ExcelDoc(path, this, doc, readOnly);
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
        public ExcelDoc TryOpenDoc(string path, bool readOnly)
        {
            Workbook doc = null;
            try
            {
                if (App == null) throw new Exception($"Не удалось открыть документ {path} приложение не запущено");
                if (!readOnly && IsDocumentOpen(path)) throw new Exception($"Документ уже редактируется {path}");
                TryPath = path;
                doc = App.Workbooks.Open(path, Type.Missing, readOnly);
                ExcelDoc d = new ExcelDoc(path, this, doc, readOnly);
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
        /// Close document
        /// - Закрыть документ
        /// </summary>
        /// <param name="path"></param>
        internal void CloseDoc(string path)
        {
            if (Docs.FirstOrDefault(x => x.Path == path) is ExcelDoc d)
            {
                if (d.IsBlock)
                    d.Close(false);
                else
                    Docs.Remove(d);
            }
        }

        /// <summary>
        /// Close
        /// - Закрыть приложение
        /// </summary>
        [HandleProcessCorruptedStateExceptions]
        public void Close()
        {
            try
            {
                for (int i = 0; i < Docs.Count; i++)
                    Docs[i].Close(false);
                App.WorkbookOpen -= App_WorkbookOpen;
                App.WorkbookBeforeClose -= App_WorkbookBeforeClose;
                App.Quit();
                Marshal.FinalReleaseComObject(App);
                Thread.Sleep(300);
            }
            catch (Exception) { }
            Clear();
        }

        /// <summary>
        /// Close
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

                if (!IsRegistred("EXCEL.EXE")) throw new Exception("Excel не установлен");
                int[] ps1 = Process.GetProcessesByName("Excel").Select(x => x.Id).ToArray();

                App = new Application();
                App.Visible = visible;
                App.DisplayAlerts = false;

                int[] ps2 = Process.GetProcessesByName("Excel").Select(x => x.Id).ToArray();
                for (int i = 0; i < ps2.Length; i++)
                    if (!ps1.Contains(ps2[i]))
                    {
                        Process p = Process.GetProcessById(ps2[i]);
                        if (p.MainWindowHandle.ToInt32() == 0)
                            AppProcess.Add(p);
                    }

                App.WorkbookOpen += App_WorkbookOpen;
                App.WorkbookBeforeClose += App_WorkbookBeforeClose;
            }
            catch (Exception e)
            {
                Close();
                throw e;
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
                    for (int j = 1; j <= Apps[i].Workbooks.Count; j++)
                        if (Path.Combine(Apps[i].Workbooks[j].Path, Apps[i].Workbooks[j].Name) == path)
                            return true;
            }
            catch (Exception) { }
            return false;
        }

        private void App_WorkbookBeforeClose(Workbook Wb, ref bool Cancel)
        {
            string dpath = Path.Combine(Wb.Path, Wb.Name);
            if (Docs.FirstOrDefault(x => x.IsBlock && x.Path == dpath) != null)
                Cancel = true;
        }

        [HandleProcessCorruptedStateExceptions]
        private void App_WorkbookOpen(Workbook Wb)
        {
            try
            {
                string dpath = Path.Combine(Wb.Path, Wb.Name);
                bool readOnly = Wb.ReadOnly;
                App.Visible = Visible;
                if (TryPath != dpath)
                {
                    Wb.Close();
                    Application App2 = new Application();
                    App2.Workbooks.Open(dpath, Type.Missing, readOnly);
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
