using System;
using Table;

namespace Teste
{
    public class Program
    {
        public static TreatsData DataT;
        public static ValidateData Validate;
        public static void Main(string[] args)
        {
            Console.Write("Tratamento de tabela: ");
            try
            {
                //aqui vc passa o local do arquivo e abre a conecao com a tabela
                DataT = new TreatsData(@"Base de Dados.xlsx");


                DataT.Clear();
                //metodo q elimina duplicidade
                DataT.DuplicityEliminate();

                //metodo elimina duplicidades sem logica 
                DataT.mergeTable("Qual a media da sua renda? (Em Salários Mínimos)",
                    "Incluindo você, qual a soma da renda das pessoas que residem com você?",
                    "Qual a media da sua renda? (Em Salários Mínimos)");

                DataT.mergeTable("Como você respondeu na pergunta acima que estuda no periodo matutino, em qual período exerce atividade remunerada:",
                   "Como você respondeu na pergunta acima que estuda no periodo noturno, em qual período exerce atividade remunerada:",
                   "Em qual período exerce atividade remunerada:");
                //metodo q preenche os campos vazios com null
                DataT.TreatsNull();


                //salva as alteracoes 
                DataT.Save();

                DataT.CloseExcel();

                Console.WriteLine("Confere");
            }
            catch
            {
                Console.WriteLine("Erro");
            }
            Console.Write("Validação: ");
            try
            {
                //fecha a a conexao

                Validate = new ValidateData(@"Base de Dados.xlsx");

                Validate.validateEmail();

                Validate.validateIdade();

                Validate.validateDuplic();

                Validate.validateAll();

                var t = Validate.CompareTable(@"Validate.xlsx");

                Validate.Save();

                Validate.Close();
                if (!(t))
                    throw new Exception();
                Console.WriteLine("Confere");
            }
            catch
            {
                Console.WriteLine("Erro");
            }
            var Convert = new ConvertData(@"Base de Dados.xlsx");

            Convert.CreateJson();

            Console.WriteLine("Aperte uma tecla para finalizar ");
            Console.ReadKey();
        }
    }
}
