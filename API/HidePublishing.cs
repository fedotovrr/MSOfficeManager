using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace MSOfficeManager.API
{
    /// <summary>
    /// Перехват и скрытие окна публикации в MS Office
    /// </summary>
    internal class HidePublishing
    {
        /// <summary>
        /// Поток поиска
        /// </summary>
        private Thread HP;

        /// <summary>
        /// Переменная действия
        /// </summary>
        private bool Do;


        /// <summary>
        /// Запуск
        /// </summary>
        public void Start()
        {
            Do = true;
            HP = new Thread(Process);
            HP.Name = "HidePublishing";
            HP.Start();
        }

        /// <summary>
        /// Останока
        /// </summary>
        public void Stop()
        {
            Do = false;
            HP.Join();
            HP = null;
        }

        /// <summary>
        /// Метод поиска
        /// </summary>
        private void Process()
        {
            while (Do)
            {
                IntPtr hWnd = new IntPtr();
                hWnd = FindWindow(null, "Публикация...");
                if (hWnd == new IntPtr())
                    hWnd = FindWindow(null, "Publishing...");
                if (hWnd != new IntPtr())
                    ShowWindow(hWnd, 6);
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int ShowWindow(IntPtr Hwnd, int state);
    }
}
