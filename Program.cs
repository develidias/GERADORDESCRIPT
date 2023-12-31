using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using System.IO;
using Xceed.Document.NET;
using Xceed.Words.NET;

class Pessoa
{
    public string Nome { get; set; }
    public string CPF { get; set; }
    public string DataNascimento { get; set; }
    public string Banco { get; set; }
    public string Agencia { get; set; }
    public string Conta { get; set; }
    public string Telefone { get; set; }
    public List<Contrato> Contratos { get; set; }
}

class Contrato
{
    public string TipoDeContrato { get; set; }
    public decimal ValorDeContrato { get; set; }
    public decimal ValorDaParcela { get; set; }
}

class Program
{
    static void Main(string[] args)
    {
        var pessoas = new Dictionary<string, Pessoa>();

        using (var package = new ExcelPackage(new FileInfo("C:/Users/Eliesio/Desktop/TESTE.xlsx")))
        {
            var worksheet = package.Workbook.Worksheets[0]; // A primeira planilha

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                var cpf = worksheet.Cells[row, 3].Value.ToString(); // CPF está na coluna 3

                if (!pessoas.ContainsKey(cpf))
                {
                    pessoas[cpf] = new Pessoa
                    {
                        Nome = worksheet.Cells[row, 2].Value.ToString(), // Nome está na coluna 2
                        CPF = cpf,
                        DataNascimento = worksheet.Cells[row, 4].Value.ToString(),
                        Banco = worksheet.Cells[row, 7].Value.ToString(), // Banco está na coluna 7
                        Agencia = worksheet.Cells[row, 8].Value.ToString(), // Agência está na coluna 8
                        Conta = worksheet.Cells[row, 9].Value.ToString(), // Conta está na coluna 9
                        // Telefone = worksheet.Cells[row, 10].Value.ToString(), // Telefone está na coluna 10
                        Contratos = new List<Contrato>()
                    };
                }

                var contrato = new Contrato
                {
                    TipoDeContrato = worksheet.Cells[row, 1].Value.ToString(), // Tipo de Contrato está na coluna 11
                    ValorDeContrato = decimal.Parse(worksheet.Cells[row, 5].Value.ToString()), // Valor Contratos está na coluna 5
                    ValorDaParcela = decimal.Parse(worksheet.Cells[row, 6].Value.ToString()) // Valor Parcela está na coluna 6
                };

                pessoas[cpf].Contratos.Add(contrato);
            }
        }

        foreach (var pessoa in pessoas.Values)
        {
            var doc = DocX.Create($"{pessoa.Nome} - SCRIPT.docx");

            doc.InsertParagraph($"Olá, Boa tarde gostaria de falar por gentileza com o (a) Sr. (a) {pessoa.Nome},");
            doc.InsertParagraph("meu nome é FORMALIZADOR,");
            doc.InsertParagraph("Falo do setor de pós vendas vou estar realizando agora algumas confirmações referente a proposta do crédito consignado que o senhor contratou Ok?");
            doc.InsertParagraph($"NOME COMPLETO: {pessoa.Nome}");
            doc.InsertParagraph($"CPF: {pessoa.CPF}");
            doc.InsertParagraph($"DATA DE NASCIMENTO: {pessoa.DataNascimento}");

            var valorTotalContratos = pessoa.Contratos.Sum(c => c.ValorDeContrato);
            var valorTotalParcelas = pessoa.Contratos.Sum(c => c.ValorDaParcela);

            doc.InsertParagraph($"Verifiquei aqui no sistema que o senhor contratou o valor de R$ {valorTotalContratos} em 84x de R$ {valorTotalParcelas} fracionado nos seguintes contratos:");
            foreach (var contrato in pessoa.Contratos)
            {
                doc.InsertParagraph($"1 {contrato.TipoDeContrato} de R$ {contrato.ValorDeContrato} parcela de R$ {contrato.ValorDaParcela} – 84x");
            }

            doc.InsertParagraph($"E o valor foi depositado no Banco: {pessoa.Banco} Ag: {pessoa.Agencia} C/C: {pessoa.Conta}");
            doc.InsertParagraph("O Sr(a) confirma esta contratação?");
            doc.InsertParagraph("R: Confirmo");
            doc.InsertParagraph("Certo, muito obrigado pelas confirmações vou estar anexando a gravação em nosso sistema.");
            doc.InsertParagraph("Vou passar algumas instruções para o senhor, que é para o senhor(a) ficar atento caso venha acontecer com você, se entrarem em contato com o senhor(a) após o valor ser liberado em conta solicitando devoluções do valor através de PIX, TRANSFERENCIA, PAGAMENTO DE BOLETO não é para o senhor(a) realizar esse tipo de procedimento pois não faz parte da nossa empresa e nem da nossa maneira de trabalho, além do mais isso é um golpe... fique atenta, se acontecer, o senhor(a) tem o contato da operadora que fechou a proposta, entre em contato imediatamente, tem alguma dúvida ?");
            doc.InsertParagraph("R:");
            doc.InsertParagraph("Ok, só lembrando que essa ligação fica anexada em nosso sistema caso isso venha ocorrer com a senhora não nos responsabilizamos, pois, a instrução foi passada acima.");
            doc.InsertParagraph("O senhor está de acordo? R:  SIM");
            doc.InsertParagraph("Ok, muito obrigada pela atenção, Boa tarde!!");
            doc.InsertParagraph($"TELEFONE: {pessoa.Telefone}");

            doc.Save();
        }
    }
}
