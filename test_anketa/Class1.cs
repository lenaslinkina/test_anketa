using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;


namespace test_anketa
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel()
        {


        }

        //static void Main(string[] args)
        //{

        //    wb.ChartObjects chartObjs = (wb.ChartObjects)workSheet.ChartObjects();
        //    Excel.ChartObject chartObj = chartObjs.Add(5, 50, 300, 300);
        //}
        public Excel(string path, int Sheet)
        {
            this.path = path;
            // excel = new _Excel.Application();
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];

        }
        public void CreateNewFile()
        {
            this.wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            // this.ws = wb.Worksheets[1];
        }

        public void CreateNewSheet()
        {
            Worksheet temptsheet = wb.Worksheets.Add(After: ws);
        }
        public string ReadCell(int i, int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].Value2 != null)
                return Convert.ToString(ws.Cells[i, j].Value2);
            else
                return "";




        }
        public void range (string s1, string s2)
        {
            var cells = ws.get_Range(s1, s2);
            cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние вертикальные
            cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние горизонтальные            
            cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
            cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
            cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
          cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
        }

       
        
    


        public void WriteToCell (int i, int j, string s)
        {
            i++;
            j++;
            ws.Cells[i, j].Value2 = s;
        }
        public void addchart()

        {

            //// Вычисляем сумму этих чисел
            //_Excel.Range rng = workSheet.Range["A2"];
            //rng.Formula = "=SUM(A1:L1)";
            //rng.FormulaHidden = false;
           
           // excel.ActiveSheet.Range["A2:C2"].Borders.
            //Выделяем границы у этой ячейки
            //_Excel.Borders border = rng.Borders;
            //border.LineStyle = _Excel.XlLineStyle.xlContinuous;

            // Строим  диаграмму
            if (ws.ChartObjects().Count != 0)
            {
                excel.Visible = true;
                excel.UserControl = true;

            }
            else
            {
                _Excel.ChartObjects chartObjs = (_Excel.ChartObjects)ws.ChartObjects();
                _Excel.ChartObject chartObj = chartObjs.Add(200, 50, 300, 200);
                _Excel.Chart xlChart = chartObj.Chart;
                _Excel.Range rng = ws.Range["B2:B8", "C2:C8"];
                // Устанавливаем тип диаграммы
                xlChart.ChartType = _Excel.XlChartType.xlColumnStacked;
                // Устанавливаем источник данных 
                xlChart.SetSourceData(rng);


                // Открываем созданный excel-файл
                excel.Visible = true;
                excel.UserControl = true;



            }


        }

        public double CellNull(Excel excel, int i, int j)
        {
            i++;
            j++;
            double a = 0;
            if (ws.Cells[i, j].Value2 != null)
            {
                 a = Convert.ToDouble(ws.Cells[i, j].Value2);
                return a;
            }
                    
            else
            {
                a = 0;
                return a ;
            }
        }
        public void Save()
        {
            wb.Save();
        }
        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }

        public void Close()
        {
            wb.Close();
        }
    }
}
