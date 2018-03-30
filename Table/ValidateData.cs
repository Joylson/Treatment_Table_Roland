using System;
using OfficeOpenXml;
using System.IO;
using System.Text.RegularExpressions;

namespace Table
{
    public class ValidateData
    {
        private FileInfo _File = null;
        private ExcelPackage _Excel = null;
        private ExcelWorksheet _Sheet = null;
        private int ColValidate = 0;
        private int Row = 0;
        private int Col = 0;

        public ValidateData(string file)
        {
            _File = new FileInfo(file);
            _Excel = new ExcelPackage(_File);

            _Sheet = _Excel.Workbook.Worksheets[2];


            while (_Sheet.Cells[1, Col + 1].Value != null)
            {
                Col++;

            }
            while (_Sheet.Cells[Row + 1, 1].Value != null)
            {
                Row++;
            }

            _Sheet.Cells[1, Col + 1].Value = "Validade";

            ColValidate = Col + 1;

        }

        public void validateEmail()
        {
            int colEmail = 0;
            Regex reg = new Regex(@"^[A-Za-z0-9](([_\.\-]?[a-zA-Z0-9]+)*)@([A-Za-z0-9]+)(([\.\-]?[a-zA-Z0-9]+)*)\.([A-Za-z]{2,})$");

            for (int i = _Sheet.Dimension.Start.Column; i <= _Sheet.Dimension.End.Column; i++)
            {
                if (_Sheet.Cells[1, i].Value != null && _Sheet.Cells[1, i].Value.ToString() == "Endereço de e-mail")
                {
                    colEmail = i;
                }
            }

            for (int i = 2; i < Row; i++)
            {
                if (_Sheet.Cells[i, colEmail].Value != null)
                {
                    if (!reg.IsMatch(_Sheet.Cells[i, colEmail].Value.ToString()) && _Sheet.Cells[i, ColValidate].Value.ToString() != "Invalido")
                    {
                        _Sheet.Cells[1, ColValidate].Value = "Invalido";
                    }
                }
            }
        }

        public void validateIdade()
        {
            int colIdade = 0;

            for (int i = _Sheet.Dimension.Start.Column; i <= _Sheet.Dimension.End.Column; i++)
            {
                if (_Sheet.Cells[1, i].Value != null && _Sheet.Cells[1, i].Value.ToString() == "Idade:")
                {
                    colIdade = i;
                }
            }

            for (int i = 2; i < Row; i++)
            {
                if (_Sheet.Cells[i, colIdade].Value != null)
                {
                    String idade = DateTime.FromOADate(Convert.ToDouble(_Sheet.Cells[i, colIdade].Value)).ToString().Split('/', ' ')[2];
                    if (Convert.ToInt32(idade) >= 2000)
                    {
                        _Sheet.Cells[i, ColValidate].Value = "Invalido";
                    }
                }
            }
        }

        public void validateDuplic()
        {
            int colName = 0, colRa = 0;

            for (int i = 1; i <= Col; i++)
            {
                if (_Sheet.Cells[1, i].Value.ToString() == "Nome:")
                {
                    colName = i;
                }
                else if (_Sheet.Cells[1, i].Value.ToString() == "RA:")
                {
                    colRa = i;
                }
            }

            for (int i = 2; i <= Row; i++)
            {

                for (int c = 2; c <= Row; c++)
                {
                    if (c != i && _Sheet.Cells[c, colRa].Value.ToString().Replace(" ", "") == _Sheet.Cells[i, colRa].Value.ToString().Replace(" ", "") && _Sheet.Cells[c, colName].Value.ToString().Replace(" ", "") == _Sheet.Cells[i, colName].Value.ToString().Replace(" ", ""))
                    {
                        _Sheet.Cells[i, ColValidate].Value = "Invalido";
                    }
                }
            }
        }

        public void validateAll()
        {

            for (int i = 2; i <= Row; i++)
            {
                if (_Sheet.Cells[i, ColValidate].Value == null || _Sheet.Cells[i, ColValidate].Value.ToString() != "Invalido")
                {
                    _Sheet.Cells[i, ColValidate].Value = "Valido";
                }
            }
        }

        public Boolean CompareTable(string table)
        {
            ExcelPackage excel = new ExcelPackage(new FileInfo(table));

            var sheet = excel.Workbook.Worksheets[2];
            var colSheet1 = 0;
            var colSheet2 = 0;

            while (sheet.Cells[1,colSheet1 + 1].Value != null) { colSheet1++; }
            while (_Sheet.Cells[1, colSheet2 + 1].Value != null) { colSheet2++; }

            if (colSheet1 != colSheet2)
            {
                return false;
            }

            for (int i = 1; i <= sheet.Dimension.End.Column; i++)
            {
                var check = false;
                if (sheet.Cells[1,i].Value != null)
                {
                    for (int c = 1; c <= _Sheet.Dimension.End.Column; c++)
                    {
                        if (_Sheet.Cells[1,c].Value != null && _Sheet.Cells[1,c].Value.ToString() == sheet.Cells[1,i].Value.ToString())
                        {
                            check = true;
                        }
                    }
                    if (!check)
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public void Save()
        {
            _Excel.SaveAs(_File);
        }
        public void Close()
        {
            _Excel.Dispose();
        }
    }
}
