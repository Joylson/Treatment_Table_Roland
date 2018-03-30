using System;
using System.Runtime.Serialization;

namespace Table
{
    [DataContract]
    public class Data
    {
        [DataMember]
        public string RA { get; set; }
        [DataMember]
        public string Nome { get; set; }
        [DataMember]
        public string Email { get; set; }
        [DataMember]
        public string Carimbo { get; set; }
        [DataMember]
        public string Nascimento { get; set; }
        [DataMember]
        public string Deficiencia { get; set; }
        [DataMember]
        public string EstadoCivil { get; set; }
        [DataMember]
        public string Filhos { get; set; }
        [DataMember]
        public string Cidade { get; set; }
        [DataMember]
        public string Locomocao { get; set; }
        [DataMember]
        public string SituacaoDomiciliar { get; set; }
        [DataMember]
        public string TempoMoradia { get; set; }
        [DataMember]
        public string MoraCom { get; set; }
        [DataMember]
        public string Trabalha { get; set; }
        [DataMember]
        public string MediaRenda { get; set; }
        [DataMember]
        public string PeriodoEstudo { get; set; }
        [DataMember]
        public string PessoasResidem { get; set; }
        [DataMember]
        public string PessoasTrabalham { get; set; }
        [DataMember]
        public string PeriodoTrabalho { get; set; }
        [DataMember]
        public string VidaEscolar { get; set; }
        [DataMember]
        public string ConhecimentoInformatica { get; set; }
        [DataMember]
        public string MotivoVestibular { get; set; }
        [DataMember]
        public string ConhecimentoLingua { get; set; }
        [DataMember]
        public string Linguas { get; set; }
        [DataMember]
        public string Meio { get; set; }
        [DataMember]
        public string Validade { get; set; }

        public Data()
        {
            RA = "RA:";
            Nome = "Nome:";
            Email = "Endereço de e-mail";
            Carimbo = "Carimbo de data/hora";
            Nascimento = "Idade:";
            Deficiencia = "Você tem alguma deficiência(física/mental)?";
            EstadoCivil = "Estado Civil:";
            Filhos = "Filhos";
            Cidade = "Cidade em que reside:";
            Locomocao = "Meio de Locomoção:";
            SituacaoDomiciliar = "Situação Domiciliar:";
            TempoMoradia = "Tempo de Moradia(Anos):";
            MoraCom = "Com quem mora:";
            MediaRenda = "Qual a media da sua renda? (Em Salários Mínimos)";
            Linguas = "Qual(is) língua(s)?";
            Trabalha = "Exerce atividade remunerada?";
            PeriodoEstudo = "Periodo:";
            PessoasResidem = "Quantas pessoas residem com você? (Incluindo você)";
            PessoasTrabalham = "Incluindo você, quantas pessoas da sua residencia exercem atividade remunerada?";
            PeriodoTrabalho = "Em qual período exerce atividade remunerada:";
            VidaEscolar = "Sua vida escolar foi:";
            ConhecimentoInformatica = "Tem algum conhecimento em informática?";
            MotivoVestibular = "Por qual(ais) motivo(s) abaixo resolveu prestar vestibular para FATEC Franca?";
            ConhecimentoLingua = "Tem conhecimento sobre alguma língua estrangeira?";
            Meio = "Para responder a esse questionário, qual meio usou?";
            Validade = "Validade";
        }
    }
}
