using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTest
{
    // Класс  документа Excel скрывает плодробности работы с эксел, требует добавления в References библиотеки Microsoft.Office.Interop.Excel
    public class ExcelDocument
    {
        private Excel.Application _application = null; 
        private Excel.Workbook _workBook = null;
        private Excel.Worksheet _workSheet = null;
        private object _missingObj = System.Reflection.Missing.Value;

        //КОНСТРУКТОР
        public ExcelDocument()
        {
            _application = new Excel.ApplicationClass();
            _workBook = _application.Workbooks.Add(_missingObj);
            _workSheet = (Excel.Worksheet)_workBook.Worksheets.get_Item(1);
        }

        public ExcelDocument(string pathToTemplate)
        {
            object pathToTemplateObj = pathToTemplate;

            _application = new Excel.ApplicationClass();
            _workBook = _application.Workbooks.Add(pathToTemplateObj);
            _workSheet = (Excel.Worksheet)_workBook.Worksheets.get_Item(1);
        }

        // ВИДИМОСТЬ ДОКУМЕНТА
        public bool Visible
        {
            get
            {
                return _application.Visible;
             }
            set
            {
                _application.Visible = value;            
            }
        }

        public int RowsWithDataCount
        {
            get
            {
                int usedRowsNum = 0;
                usedRowsNum = _workSheet.UsedRange.Rows.Count;
                return usedRowsNum;
            }
        }

        public int ColumnsWithDataCount
        {
            get
            {
                int usedColumnsNum = 0;
                usedColumnsNum = _workSheet.UsedRange.Columns.Count;
                return usedColumnsNum;
            }
        }

        // ВСТАВКА ЗНАЧЕНИЯ В ЯЧЕЙКУ
        public void SetCellValue(string cellValue, int rowIndex, int columnIndex)
        {
            _workSheet.Cells[rowIndex, columnIndex] = cellValue;
        }

        // ЧТЕНИЕ ЗНАЧЕНИЯ
        public string GetCellValue(int rowIndex, int columnIndex)
        {
            string cellValue = "";
            
            
            Excel.Range cellRange = (Excel.Range) _workSheet.Cells[rowIndex, columnIndex];
            if (cellRange.Value != null)
            {
                cellValue = cellRange.Value.ToString();
            }
            return cellValue;
        }

        public void Close()
        {
            _workBook.Close(false, _missingObj, _missingObj);

            _application.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(_application);
           
            _application = null;
            _workBook = null;
            _workSheet = null;

            System.GC.Collect();
        }
    }
}

