using System;
using System.Linq;

namespace MSOfficeManager.Templates
{
    /// <summary>
    /// Шаблон заполнения таблицы
    /// </summary>
    public class TemplateTable
    {
        /// <summary>
        /// Template header table - Шаблон заголовка таблицы
        /// </summary>
        public HeaderCell[] Header;

        /// <summary>
        /// Read start row number - Строка начала чтения таблицы
        /// </summary>
        public int StartRow = 1;

        /// <summary>
        /// Value if the table is found - Была ли найдена таблица с данным шаблоном
        /// </summary>
        public bool Found = false;

        /// <summary>
        /// Table header equivalence value - Значение эквивалентности заголовка таблицы
        /// </summary>
        /// <param name="header">заголовок таблицы</param>
        /// <returns></returns>
        public bool IsEqual(string[] header)
        {
            if (header == null || Header == null)
                return false;
            for (int j = 0; j < Header.Length; j++)
                Header[j].ClearState();
            for (int i = 0; i < header.Length; i++)
                for (int j = 0; j < Header.Length; j++)
                    if (Header[j].Identify != null && Header[j].Identify(header[i]))
                    {
                        Header[j].SetState(i);
                        break;
                    }
            for (int i = 0; i < Header.Length; i++)
                if (Header[i].IsRequired && !Header[i].Found)
                {
                    for (int j = 0; j < Header.Length; j++)
                        Header[j].ClearState();
                    return false;
                }
            Found = true;
            return true;
        }

        /// <summary>
        /// Get table content - Получить контент таблицы
        /// </summary>
        /// <returns></returns>
        public virtual Cell[][] GetCells()
        {
            return null;
        }

        /// <summary>
        /// Row to object - Конвертирование строки в объект
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public virtual object ToItem(string[] row)
        {
            return row == null ? null : Header.Select(x => GetValueByHeader(row, x.Column)).ToArray();
        }

        /// <summary>
        /// Get value by header index - Получить значение для столбца по индексу заголовка
        /// </summary>
        /// <param name="row"></param>
        /// <param name="headerIndex"></param>
        /// <returns></returns>
        public string GetValueByHeader(string[] row, int headerIndex)
        {
            return Header != null && headerIndex < Header.Length && 
                Header[headerIndex].Found && Header[headerIndex].Column > -1 && Header[headerIndex].Column < row.Length ? 
                row[Header[headerIndex].Column] : null;
        }
    }

    /// <summary>
    /// Ячейка шаблона таблицы
    /// </summary>
    public class Cell
    {
        /// <summary>
        /// Значение ячейки
        /// </summary>
        public string Value;

        /// <summary>
        /// Столбец ячейки
        /// </summary>
        public int Column = -1;

        /// <summary>
        /// Протяженность ячейки
        /// </summary>
        public int ColumnSpan = 1;

        public Cell(int column, string value)
        {
            Value = value;
            Column = column;
            ColumnSpan = 1;
        }

        public Cell(string value = null, int column = -1, int columnSpan = 1)
        {
            Value = value;
            Column = column;
            ColumnSpan = columnSpan;
        }

        internal int LastMergeColumn
        {
            get
            {
                int c = Column + ColumnSpan - 1;
                if (c > Column)
                    return c;
                return Column;
            }
        }
    }

    /// <summary>
    /// Ячейка шаблона заголовка таблицы
    /// </summary>
    public class HeaderCell
    {
        /// <summary>
        /// Function identify - Функция идентификации
        /// </summary>
        /// <returns></returns>
        public Func<string, bool> Identify { get; private set; }

        /// <summary>
        /// Is presence required - Статус обязательного присутсвия
        /// </summary>
        public bool IsRequired { get; private set; }

        /// <summary>
        /// Value if the cell header is found - Была ли найдена ячейка с данным шаблоном
        /// </summary>
        public bool Found { get; private set; }

        /// <summary>
        /// Column number header cell - Столбец ячейки
        /// </summary>
        public int Column { get; private set; } = -1;


        /// <summary>
        /// Ячейка шаблона заголовка таблицы
        /// </summary>
        /// <param name="isRequired">статус обязательного присутсвия</param>
        /// <param name="identify">функция идентификации</param>
        public HeaderCell(bool isRequired, Func<string, bool> identify)
        {
            IsRequired = isRequired;
            Identify = identify;
        }

        internal void SetState(int column)
        {
            if (Found || column < 0) return;
            Found = true;
            Column = column;
        }

        internal void ClearState()
        {
            Found = false;
            Column = -1;
        }
    }
}
