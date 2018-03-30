using System;
using System.IO;
using OfficeOpenXml;
using System.Collections.Generic;

namespace Table
{
    public class TreatsData
    {
        private ExcelPackage _Excel = null;
        private FileInfo _File;
        private ExcelWorksheet tableChange = null;
        private ExcelWorksheet tableOriginal = null;

        public TreatsData(string table)
        {
            _File = new FileInfo(table);
            _Excel = new ExcelPackage(new FileInfo(table));
            tableOriginal =  _Excel.Workbook.Worksheets[1];
            tableOriginal.Name = "Original";
            try
            {
                tableChange = _Excel.Workbook.Worksheets.Add("Change");
            }
            catch
            {
                tableChange = _Excel.Workbook.Worksheets[2];
            }
        }


        public void Clear()
        {
            _Excel.Workbook.Worksheets.Delete(tableChange);
            tableChange = _Excel.Workbook.Worksheets.Add("Change");
        }

        public void DuplicityEliminate()
        {
            //lista com todos os titulos
            List<string> titles = new List<string>();
            //Checa todos os titulos 
            for (int i = tableOriginal.Dimension.Start.Column; i <= tableOriginal.Dimension.End.Column; i++)
            {
                //verifica se eles sao nulos se não for prosegue 
                if (tableOriginal.Cells[1, i].Value != null)
                {
                    //checa se ja existe o titulo na lista se n existe adiciona
                    var equality = false;
                    foreach (string title in titles)
                    {
                        if (tableOriginal.Cells[1, i].Value.ToString() == title)
                        {
                            equality = true;
                        }
                    }
                    if (!(equality))
                    {
                        titles.Add(tableOriginal.Cells[1, i].Value.ToString());
                    }
                }
            }
            //adiciona todos os titulos na tabela de alteração
            var col = 1;
            foreach (string title in titles)
            {
                tableChange.Cells[1, col].Value = title;
                col++;
            }
            agroupData();
        }

        public void SaveAS()
        {

        }

        private void agroupData()
        {
            //pega todos os titulos ja adicionados na tabela de alteracao 
            for (int i = tableChange.Dimension.Start.Column; i <= tableChange.Dimension.End.Column; i++)
            {
                //checa se e nulo
                if (tableChange.Cells[1, i].Value != null)
                {
                    //pega todos os titulos na tabela original
                    for (int c = tableOriginal.Dimension.Start.Column; c <= tableOriginal.Dimension.End.Column; c++)
                    {
                        //checa se o titulo da tabela original e igual ao de alteracao
                        if (tableChange.Cells[1, i].Value == tableOriginal.Cells[1, c].Value && tableOriginal.Cells[1, c].Value != null)
                        {
                            //adiciona linha por linha da coluna correspondente ao titulo
                            for (int x = 2; x <= tableOriginal.Dimension.End.Row; x++)
                            {
                                if (tableOriginal.Cells[x, c].Value != null)
                                {
                                    tableChange.Cells[x, i].Value = tableOriginal.Cells[x, c].Value;
                                    tableChange.Cells[x, i].Style.Numberformat.Format = "General";
                                }
                            }
                        }
                    }
                }
            }
        }

        public void TreatsNull()
        {
            //percorre a tabela de alteracao
            for (int i = tableChange.Dimension.Start.Column; i <= tableChange.Dimension.End.Column; i++)
            {
                //checa se a campos vazios 
                if (tableChange.Cells[1, i].Value != null)
                {
                    for (int c = 2; c <= tableChange.Dimension.End.Row; c++)
                    {
                        if (tableChange.Cells[c, i].Value == null)
                        {
                            tableChange.Cells[c, i].Value = "null";
                        }
                    }
                }
            }
        }

        public void mergeTable(string title1, string title2, string titleOriginal)
        {
            int column1 = 0;
            int column2 = 0;
            for (int i = tableChange.Dimension.Start.Column; i <= tableChange.Dimension.End.Column; i++)
            {
                if (tableChange.Cells[1, i].Value != null)
                {
                    if (tableChange.Cells[1, i].Value.ToString() == title1)
                    {
                        column1 = i;
                    }
                    else if (tableChange.Cells[1, i].Value.ToString() == title2)
                    {
                        column2 = i;
                    }
                }
            }

            var col = 0;
            while (tableChange.Cells[1,col + 1].Value != null)
            {
                col++;
            }

            if (column1 != 0 && column2 != 0)
            {
                tableChange.Cells[1, col + 1].Value = titleOriginal;

                for (int i = 2; i <= tableChange.Dimension.End.Row; i++)
                {
                    if (tableChange.Cells[i, column1].Value != null)
                    {
                        tableChange.Cells[i, col + 1].Value = tableChange.Cells[i, column1].Value;
                    }
                    if (tableChange.Cells[i, column2].Value != null)
                    {
                        tableChange.Cells[i, col + 1].Value = tableChange.Cells[i, column2].Value;
                    }
                }
                if(column1 > column2)
                {
                    tableChange.DeleteColumn(column1);
                    tableChange.DeleteColumn(column2);
                }
                else
                {
                    tableChange.DeleteColumn(column2);
                    tableChange.DeleteColumn(column1);
                }
            }
        }

       

        public ExcelPackage GetPackage()
        {
            return _Excel = new ExcelPackage(_File);
        }


        public void Save()
        {
            _Excel.SaveAs(_File);
        }

        public void CloseExcel()
        {
            _Excel.Dispose();
        }
    }
}
