using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace JiangLiQuery.ExcelHelper
{
    public class ExcelRender
    {
        public ExcelRender() {
            
        }

        /// <summary>
        /// 根据Excel列类型获取列的值
        /// </summary>
        /// <param name="cell">Excel列</param>
        /// <returns></returns>
        private string GetCellValue(ICell cell)
        {
            if (cell == null)
                return string.Empty;
            switch (cell.CellType)
            {
                case CellType.Blank:
                    return string.Empty;
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    return cell.ErrorCellValue.ToString();
                case CellType.Numeric:
                case CellType.Unknown:
                default:
                    return cell.ToString();//This is a trick to get the correct value of the cell. NumericCellValue will return a numeric value no matter the cell value is a date or a number
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Formula:
                    try
                    {
                        HSSFFormulaEvaluator e = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                        e.EvaluateInCell(cell);
                        return cell.ToString();
                    }
                    catch
                    {
                        return cell.NumericCellValue.ToString();
                    }
            }
        }

        /// <summary>
        /// 自动设置Excel列宽
        /// </summary>
        /// <param name="sheet">Excel表</param>
        private void AutoSizeColumns(ISheet sheet)
        {
            if (sheet.PhysicalNumberOfRows > 0)
            {
                IRow headerRow = sheet.GetRow(0);

                for (int i = 0, l = headerRow.LastCellNum; i < l; i++)
                {
                    sheet.AutoSizeColumn(i);
                }
            }
        }

        /// <summary>
        /// 保存Excel文档流到文件
        /// </summary>
        /// <param name="ms">Excel文档流</param>
        /// <param name="fileName">文件名</param>
        private void SaveToFile(MemoryStream ms, string fileName)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                byte[] data = ms.ToArray();

                fs.Write(data, 0, data.Length);
                fs.Flush();

                data = null;
            }
        }

        /// <summary>
        /// DataReader转换成Excel文档流
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        public MemoryStream RenderToExcel(IDataReader reader)
        {
            MemoryStream ms = new MemoryStream();

            using (reader)
            {
                IWorkbook workbook = new HSSFWorkbook();
                ISheet sheet = workbook.CreateSheet();
                IRow headerRow = sheet.CreateRow(0);
                int cellCount = reader.FieldCount;

                // handling header.
                for (int i = 0; i < cellCount; i++)
                {
                    headerRow.CreateCell(i).SetCellValue(reader.GetName(i));
                }

                // handling value.
                int rowIndex = 1;
                while (reader.Read())
                {
                    IRow dataRow = sheet.CreateRow(rowIndex);

                    for (int i = 0; i < cellCount; i++)
                    {
                        dataRow.CreateCell(i).SetCellValue(reader[i].ToString());
                    }

                    rowIndex++;
                }

                AutoSizeColumns(sheet);

                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
            }
            return ms;
        }

        /// <summary>
        /// DataReader转换成Excel文档流，并保存到文件
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="fileName">保存的路径</param>
        public void RenderToExcel(IDataReader reader, string fileName)
        {
            using (MemoryStream ms = RenderToExcel(reader))
            {
                SaveToFile(ms, fileName);
            }
        }

        /// <summary>
        /// DataTable转换成Excel文档流
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public MemoryStream RenderToExcel(DataTable table)
        {
            MemoryStream ms = new MemoryStream();

            using (table)
            {
                IWorkbook workbook = new HSSFWorkbook();
                ISheet sheet = workbook.CreateSheet();
                    IRow headerRow = sheet.CreateRow(0);

                    // handling header.
                    foreach (DataColumn column in table.Columns)
                        headerRow.CreateCell(column.Ordinal).SetCellValue(column.Caption);//If Caption not set, returns the ColumnName value

                    // handling value.
                    int rowIndex = 1;

                    foreach (DataRow row in table.Rows)
                    {
                        IRow dataRow = sheet.CreateRow(rowIndex);

                        foreach (DataColumn column in table.Columns)
                        {
                            dataRow.CreateCell(column.Ordinal).SetCellValue(row[column].ToString());
                        }

                        rowIndex++;
                    }
                    AutoSizeColumns(sheet);

                    workbook.Write(ms);
                    ms.Flush();
                    ms.Position = 0;
            }
            return ms;
        }

        /// <summary>
        /// DataTable转换成Excel文档流，并保存到文件
        /// </summary>
        /// <param name="table"></param>
        /// <param name="fileName">保存的路径</param>
        public void RenderToExcel(DataTable table, string fileName)
        {
            using (MemoryStream ms = RenderToExcel(table))
            {
                SaveToFile(ms, fileName);
            }
        }

        /// <summary>
        /// Excel文档流是否有数据
        /// </summary>
        /// <param name="excelFileStream">Excel文档流</param>
        /// <returns></returns>
        public bool HasData(Stream excelFileStream)
        {
            return HasData(excelFileStream, 0);
        }

        /// <summary>
        /// Excel文档流是否有数据
        /// </summary>
        /// <param name="excelFileStream">Excel文档流</param>
        /// <param name="sheetIndex">表索引号，如第一个表为0</param>
        /// <returns></returns>
        public bool HasData(Stream excelFileStream, int sheetIndex)
        {
            using (excelFileStream)
            {
                IWorkbook workbook = new HSSFWorkbook(excelFileStream);
                    if (workbook.NumberOfSheets > 0)
                    {
                        if (sheetIndex < workbook.NumberOfSheets)
                        {
                            ISheet sheet = workbook.GetSheetAt(sheetIndex);
                            return sheet.PhysicalNumberOfRows > 0;
                            
                        }
                    }
                
            }
            return false;
        }

        /// <summary>
        /// Excel文档流转换成DataTable
        /// 第一行必须为标题行
        /// </summary>
        /// <param name="excelFileStream">Excel文档流</param>
        /// <param name="sheetName">表名称</param>
        /// <returns></returns>
        public DataTable RenderFromExcel(Stream excelFileStream, string sheetName)
        {
            return RenderFromExcel(excelFileStream, sheetName, 0);
        }

        /// <summary>
        /// Excel文档流转换成DataTable
        /// </summary>
        /// <param name="excelFileStream">Excel文档流</param>
        /// <param name="sheetName">表名称</param>
        /// <param name="headerRowIndex">标题行索引号，如第一行为0</param>
        /// <returns></returns>
        public DataTable RenderFromExcel(Stream excelFileStream, string sheetName, int headerRowIndex)
        {
            DataTable table = null;

            using (excelFileStream)
            {
                IWorkbook workbook = new HSSFWorkbook(excelFileStream);

                ISheet sheet = workbook.GetSheet(sheetName);
                    
                table = RenderFromExcel(sheet, headerRowIndex);
                    
                
            }
            return table;
        }

        /// <summary>
        /// Excel文档流转换成DataTable
        /// 默认转换Excel的第一个表
        /// 第一行必须为标题行
        /// </summary>
        /// <param name="excelFileStream">Excel文档流</param>
        /// <returns></returns>
        public DataTable RenderFromExcel(Stream excelFileStream)
        {
            return RenderFromExcel(excelFileStream, 0, 0);
        }

        /// <summary>
        /// Excel文档流转换成DataTable
        /// 第一行必须为标题行
        /// </summary>
        /// <param name="excelFileStream">Excel文档流</param>
        /// <param name="sheetIndex">表索引号，如第一个表为0</param>
        /// <returns></returns>
        public DataTable RenderFromExcel(Stream excelFileStream, int sheetIndex)
        {
            return RenderFromExcel(excelFileStream, sheetIndex, 0);
        }

        /// <summary>
        /// Excel文档流转换成DataTable
        /// </summary>
        /// <param name="excelFileStream">Excel文档流</param>
        /// <param name="sheetIndex">表索引号，如第一个表为0</param>
        /// <param name="headerRowIndex">标题行索引号，如第一行为0</param>
        /// <returns></returns>
        public DataTable RenderFromExcel(Stream excelFileStream, int sheetIndex, int headerRowIndex)
        {
            DataTable table = null;

            using (excelFileStream)
            {
                IWorkbook workbook = new HSSFWorkbook(excelFileStream);
                
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                    
                table = RenderFromExcel(sheet, headerRowIndex);
                    
                
            }
            return table;
        }

        /// <summary>
        /// Excel表格转换成DataTable
        /// </summary>
        /// <param name="sheet">表格</param>
        /// <param name="headerRowIndex">标题行索引号，如第一行为0</param>
        /// <returns></returns>
        private DataTable RenderFromExcel(ISheet sheet, int headerRowIndex)
        {
            DataTable table = new DataTable();

            IRow headerRow = sheet.GetRow(headerRowIndex);
            int cellCount = headerRow.LastCellNum;//LastCellNum = PhysicalNumberOfCells
            int rowCount = sheet.LastRowNum;//LastRowNum = PhysicalNumberOfRows - 1

            //handling header.
            for (int i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                table.Columns.Add(column);
            }

            for (int i = (sheet.FirstRowNum + 1); i <= rowCount; i++)
            {
                IRow row = sheet.GetRow(i);
                DataRow dataRow = table.NewRow();

                if (row != null)
                {
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                            dataRow[j] = GetCellValue(row.GetCell(j));
                    }
                }

                table.Rows.Add(dataRow);
            }

            return table;
        }

        /// <summary>
        /// ExcelToDataTable
        /// </summary>
        /// <param name="excelPath">路径</param>
        /// <param name="sheetName">表名</param>
        /// <returns></returns>
        public DataTable ExcelToDataTable(string excelPath, string sheetName)
        {
            return ExcelToDataTable(excelPath, sheetName, true);
        }

        /// <summary>
        /// 将Excel转换成DataTable
        /// </summary>
        /// <param name="excelPath">路径</param>
        /// <param name="sheetName">表名</param>
        /// <param name="firstRowAsHeader">是否读取第一行</param>
        /// <returns></returns>
        public DataTable ExcelToDataTable(string excelPath, string sheetName, bool firstRowAsHeader)
        {
            using (FileStream fileStream = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
            {
                HSSFWorkbook workbook = new HSSFWorkbook(fileStream);

                HSSFFormulaEvaluator evaluator = new HSSFFormulaEvaluator(workbook);

                HSSFSheet sheet = workbook.GetSheet(sheetName) as HSSFSheet;

                return ExcelToDataTable(sheet, evaluator, firstRowAsHeader);
            }
        }

        /// <summary>
        /// ExcelToDataTable
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="evaluator"></param>
        /// <param name="firstRowAsHeader"></param>
        /// <returns></returns>
        private DataTable ExcelToDataTable(HSSFSheet sheet, HSSFFormulaEvaluator evaluator, bool firstRowAsHeader)
        {
            if (firstRowAsHeader)
            {
                return ExcelToDataTableFirstRowAsHeader(sheet, evaluator);
            }
            else
            {
                return ExcelToDataTable(sheet, evaluator);
            }
        }

        /// <summary>
        /// ExcelToDataTableFirstRowAsHeader
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="evaluator"></param>
        /// <returns></returns>
        private DataTable ExcelToDataTableFirstRowAsHeader(HSSFSheet sheet, HSSFFormulaEvaluator evaluator)
        {
            using (DataTable dt = new DataTable())
            {
                HSSFRow firstRow = sheet.GetRow(0) as HSSFRow;
                int cellCount = GetCellCount(sheet);

                for (int i = 0; i < cellCount; i++)
                {
                    if (firstRow.GetCell(i) != null)
                    {
                        dt.Columns.Add(firstRow.GetCell(i).StringCellValue ?? string.Format("F{0}", i + 1), typeof(string));
                    }
                    else
                    {
                        dt.Columns.Add(string.Format("F{0}", i + 1), typeof(string));
                    }
                }

                for (int i = 1; i <= sheet.LastRowNum; i++)
                {
                    HSSFRow row = sheet.GetRow(i) as HSSFRow;
                    DataRow dr = dt.NewRow();
                    FillDataRowByHSSFRow(row, evaluator, ref dr);
                    dt.Rows.Add(dr);
                }

                dt.TableName = sheet.SheetName;
                return dt;
            }
        }

        /// <summary>
        /// 将Excel转换成DataTable
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="evaluator"></param>
        /// <returns></returns>
        private DataTable ExcelToDataTable(HSSFSheet sheet, HSSFFormulaEvaluator evaluator)
        {
            using (DataTable dt = new DataTable())
            {
                if (sheet.LastRowNum != 0)
                {
                    int cellCount = GetCellCount(sheet);

                    for (int i = 0; i < cellCount; i++)
                    {
                        dt.Columns.Add(string.Format("F{0}", i), typeof(string));
                    }

                    for (int i = 0; i < sheet.FirstRowNum; ++i)
                    {
                        DataRow dr = dt.NewRow();
                        dt.Rows.Add(dr);
                    }

                    for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
                    {
                        HSSFRow row = sheet.GetRow(i) as HSSFRow;
                        DataRow dr = dt.NewRow();
                        FillDataRowByHSSFRow(row, evaluator, ref dr);
                        dt.Rows.Add(dr);
                    }
                }

                dt.TableName = sheet.SheetName;
                return dt;
            }
        }

        private void FillDataRowByHSSFRow(HSSFRow row, HSSFFormulaEvaluator evaluator, ref DataRow dr)
        {
            if (row != null)
            {
                for (int j = 0; j < dr.Table.Columns.Count; j++)
                {
                    HSSFCell cell = row.GetCell(j) as HSSFCell;

                    if (cell != null)
                    {
                        switch (cell.CellType)
                        {
                            case CellType.Blank:
                                dr[j] = DBNull.Value;
                                break;
                            case CellType.Boolean:
                                dr[j] = cell.BooleanCellValue;
                                break;
                            case CellType.Numeric:
                                if (DateUtil.IsCellDateFormatted(cell))
                                {
                                    dr[j] = cell.DateCellValue;
                                }
                                else
                                {
                                    dr[j] = cell.NumericCellValue;
                                }
                                break;
                            case CellType.String:
                                dr[j] = cell.StringCellValue;
                                break;
                            case CellType.Error:
                                dr[j] = cell.ErrorCellValue;
                                break;
                            case CellType.Formula:
                                cell = evaluator.EvaluateInCell(cell) as HSSFCell;
                                dr[j] = cell.ToString();
                                break;
                            default:
                                throw new NotSupportedException(string.Format("Catched unhandle CellType[{0}]", cell.CellType));
                        }
                    }
                }
            }
        }

        private int GetCellCount(HSSFSheet sheet)
        {
            int firstRowNum = sheet.FirstRowNum;

            int cellCount = 0;

            for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; ++i)
            {
                HSSFRow row = sheet.GetRow(i) as HSSFRow;

                if (row != null && row.LastCellNum > cellCount)
                {
                    cellCount = row.LastCellNum;
                }
            }

            return cellCount;
        }

        #region expand

        /// <summary>
        /// 获取表工作区数量
        /// </summary>
        /// <param name="strFileName"></param>
        /// <returns></returns>
        public int GetExcelSheetsCount(string strFileName)
        {
            HSSFWorkbook hssfworkbook;
            using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                hssfworkbook = new HSSFWorkbook(file);
            }
            return hssfworkbook.NumberOfSheets;
        }

        /// <summary>
        /// 获取工作区名称
        /// </summary>
        /// <param name="strFileName"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public string GetExcelSheetsName(string strFileName, int index)
        {
            HSSFWorkbook hssfworkbook;
            using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                hssfworkbook = new HSSFWorkbook(file);
            }
            return hssfworkbook.GetNameAt(index).SheetName;
        }

        /// <summary>
        /// 转换多个工作区
        /// </summary>
        /// <param name="strFileName"></param>
        /// <returns></returns>
        public List<DataTable> ExcelToDataTableList(string strFileName)
        {
            List<DataTable> list = new List<DataTable>();
            HSSFWorkbook hssfworkbook;
            using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                hssfworkbook = new HSSFWorkbook(file);
            }

            int sheetnumber = hssfworkbook.NumberOfSheets;//获取SHEET数量
            for (int h = 0; h < sheetnumber; h++)
            {
                ISheet sheet = hssfworkbook.GetSheetAt(h);//依次获取工作区
                if (sheet != null)
                {
                    DataTable dt = new DataTable();
                    System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

                    IRow headerRow = sheet.GetRow(0);//表头
                    int cellCount = headerRow.LastCellNum;
                    //int columnsCount = 0;
                    for (int j = 0; j < cellCount; j++)
                    {

                        ICell cell = headerRow.GetCell(j);
                        if (cell != null)
                        {
                            //columnsCount++;
                            dt.Columns.Add(cell.ToString());
                        }
                    }

                    for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        DataRow dataRow = dt.NewRow();

                        if (row != null)
                        {
                            for (int j = row.FirstCellNum; j < cellCount; j++)
                            {

                                if (row.GetCell(j) != null)
                                    dataRow[j] = row.GetCell(j).ToString();
                            }
                            dt.Rows.Add(dataRow);
                        }
                    }
                    if (dt.Rows.Count != 0)
                    {
                        if (list != null)
                        {
                            list.Add(dt);

                        }
                    }
                }
            }
            return list;



        }


        /// <summary>读取excel     
        /// 默认第一行为表头，导入第一个工作表  
        /// </summary>     
        /// <param name="strFileName">excel文档路径</param>     
        /// <returns></returns>     
        public DataTable ExcelToDataTable(string strFileName)
        {
            DataTable dt = new DataTable();

            HSSFWorkbook hssfworkbook;
            using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                hssfworkbook = new HSSFWorkbook(file);
            }


            ISheet sheet = hssfworkbook.GetSheetAt(0);

            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

            IRow headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;
            string title = "";
            for (int j = 0; j < cellCount; j++)
            {

                ICell cell = headerRow.GetCell(j);
                if (cell != null)
                {
                    dt.Columns.Add(cell.ToString());
                    title += cell.ToString() + ",";
                }
            }
            System.Console.WriteLine(title);

            for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                DataRow dataRow = dt.NewRow();

                if (row != null)
                {
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                            dataRow[j] = row.GetCell(j).ToString();
                    }
                    dt.Rows.Add(dataRow);
                }
            }
            return dt;
        }

        #endregion
    }
}
