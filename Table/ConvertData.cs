using System;
using System.Collections.Generic;
using OfficeOpenXml;
using System.IO;
using System.Runtime.Serialization.Json;

namespace Table
{
    public class ConvertData
    {
        FileInfo _File = null;
        ExcelPackage _Excel = null;
        ExcelWorksheet _Sheet = null;

        public ConvertData(string table)
        {
            _File = new FileInfo(table);
            _Excel = new ExcelPackage(_File);

            _Sheet = _Excel.Workbook.Worksheets[2];
        }

        public List<Data> CreateList()
        {
            List<Data> datas = new List<Data>();
            Data checkData = new Data();
            for (int i = 2; i <= _Sheet.Dimension.End.Row; i++)
            {
                if (_Sheet.Cells[i,1].Value != null)
                {
                    datas.Add(new Data()
                    {
                        RA = _Sheet.Cells[i, Col(checkData.RA)].Value.ToString(),
                        Nome = _Sheet.Cells[i, Col(checkData.Nome)].Value.ToString(),
                        Email = _Sheet.Cells[i, Col(checkData.Email)].Value.ToString(),
                        Carimbo = DateTime.FromOADate(Convert.ToDouble(_Sheet.Cells[i, Col(checkData.Carimbo)].Value)).ToString(),
                        Nascimento = DateTime.FromOADate(Convert.ToDouble(_Sheet.Cells[i, Col(checkData.Nascimento)].Value)).ToString(),
                        Deficiencia = _Sheet.Cells[i, Col(checkData.Deficiencia)].Value.ToString(),
                        EstadoCivil = _Sheet.Cells[i, Col(checkData.EstadoCivil)].Value.ToString(),
                        Filhos = _Sheet.Cells[i, Col(checkData.Filhos)].Value.ToString(),
                        Cidade = _Sheet.Cells[i, Col(checkData.Cidade)].Value.ToString(),
                        Locomocao = _Sheet.Cells[i, Col(checkData.Locomocao)].Value.ToString(),
                        SituacaoDomiciliar = _Sheet.Cells[i, Col(checkData.SituacaoDomiciliar)].Value.ToString(),
                        TempoMoradia = _Sheet.Cells[i, Col(checkData.TempoMoradia)].Value.ToString(),
                        MoraCom = _Sheet.Cells[i, Col(checkData.MoraCom)].Value.ToString(),
                        MediaRenda = _Sheet.Cells[i, Col(checkData.MediaRenda)].Value.ToString(),
                        Linguas = _Sheet.Cells[i, Col(checkData.Linguas)].Value.ToString(),
                        Trabalha = _Sheet.Cells[i, Col(checkData.Trabalha)].Value.ToString(),
                        PeriodoEstudo = _Sheet.Cells[i, Col(checkData.PeriodoEstudo)].Value.ToString(),
                        PessoasResidem = _Sheet.Cells[i, Col(checkData.PessoasResidem)].Value.ToString(),
                        PessoasTrabalham = _Sheet.Cells[i, Col(checkData.PessoasTrabalham)].Value.ToString(),
                        PeriodoTrabalho = _Sheet.Cells[i, Col(checkData.PeriodoTrabalho)].Value.ToString(),
                        VidaEscolar = _Sheet.Cells[i, Col(checkData.VidaEscolar)].Value.ToString(),
                        ConhecimentoInformatica = _Sheet.Cells[i, Col(checkData.ConhecimentoInformatica)].Value.ToString(),
                        MotivoVestibular = _Sheet.Cells[i, Col(checkData.MotivoVestibular)].Value.ToString(),
                        ConhecimentoLingua = _Sheet.Cells[i, Col(checkData.ConhecimentoLingua)].Value.ToString(),
                        Meio = _Sheet.Cells[i, Col(checkData.Meio)].Value.ToString(),
                        Validade = _Sheet.Cells[i, Col(checkData.Validade)].Value.ToString()

                    });
                }
            }
            return datas;
        }

        public void CreateJson()
        {

            List<Data> datas = CreateList();

            //Stream de ligação de arquivo 
            FileStream stream;
            //Formato de serialização
            var sr = new DataContractJsonSerializer(typeof(List<Data>));

            if (!(File.Exists(@"Dados.json")))
            {
                //criar arquivo
                stream = new FileStream(@"Dados.json", FileMode.Create);
            }
            else
            {
                //abrir arquivo
                stream = new FileStream(@"Dados.json", FileMode.Open);
                //datas = (List<Data>)sr.ReadObject(stream);

            }
            

            try
            {
                //modificando arquivo serializado
                stream.Position = 0;
                sr.WriteObject(stream, datas);
            }
            catch 
            {
            }
            stream.Close();
        }

        public List<Data> Deserializar()
        {
            FileStream stream;
            if (File.Exists(@"Dados.json"))
            {
                //Abrir Arquivo
                stream = new FileStream(@"Dados.json", FileMode.Open);
                var sr = new DataContractJsonSerializer(typeof(List<Data>));
                //Deserializar arquivo
                var pessoas = (List<Data>)sr.ReadObject(stream);

                return pessoas;
                stream.Close();
            }
            return null;
        }

        public int Col(string title)
        {
            for (int i = 1; i <= _Sheet.Dimension.End.Column; i++)
            {
                if (_Sheet.Cells[1,i].Value != null && _Sheet.Cells[1,i].Value.ToString() == title)
                {
                    return i;
                }
            }
            return 0;
        }

        public void Close()
        {
            _Excel.Dispose();
        }


        public void Save()
        {
            _Excel.SaveAs(_File);
        }

    }
}
