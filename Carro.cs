using System;
using NetOffice.ExcelApi;
using System.IO;

public class Carro
{
    Application ex;
    private string Modelo { get; set; }
    private string Marca { get; set; }
    private string Ano { get; set; }
    private string Placa { get; set; }
    private double Preco { get; set; }
    private string Arquivo { get; set; }

    /// <summary>
    ///Construtor da classe com parametros 
    /// </summary>
    /// <param name="Arquivo"></param>
    public Carro(string Arquivo)
    {
        //Define o caminho do arquivo excel dos clientes
        this.Arquivo = Arquivo;
        //Instancia a classe Application
        ex = new Application();
    }

    /// <summary>
    /// Metodo que obtem os dados do carro
    /// </summary>
    public void ObterDados()
    {
        Console.WriteLine("Informe o nome do cliente: ");
        this.Modelo = Console.ReadLine();

        Console.WriteLine("Informe a idade do cliente: ");
        this.Marca = Console.ReadLine();

        Console.WriteLine("Informe o email do cliente: ");
        this.Ano = Console.ReadLine();

        Console.WriteLine("Informe o email do cliente: ");
        this.Preco = Convert.ToDouble(Console.ReadLine());
        
        Console.WriteLine("Informe a placa do carro: ");
        this.Placa = Console.ReadLine();
    }

    /// <summary>
    /// Salva os dados obtidos no metodo ObterDados
    /// </summary>
    public void Salvar()
    {
        //Não mostra nenhuma modal
        ex.DisplayAlerts = false;

        //verifica em qual linha deve ser salvo o cliente, cliente novo salva um novo e cliente já cadastrado altera os dados
        int linha = VerificaLinha(this.Placa);

        //Prenche os dados na linha obtida acima e nas colunas definidas
        ex.Cells[linha, 1].Value = this.Modelo;
        ex.Cells[linha, 2].Value = this.Marca;
        ex.Cells[linha, 3].Value = this.Ano;
        ex.Cells[linha, 4].Value = this.Preco;
        ex.Cells[linha, 5].Value = this.Placa;
        ex.Cells[linha, 6].Value = "A venda";

        //Salva o arquivo no local informado
        ex.ActiveWorkbook.SaveAs(this.Arquivo);
        //fecha o Application\Excel
        ex.Quit();
        //Retira da memoria
        ex.Dispose();
    }

    /// <summary>
    /// Verifica qual linha irá incluir\alterar o carro
    /// </summary>
    
    /// <param name="placa">Placa do carro, será único</param>
    /// <returns>Linha na qual irá incluir\alterar o carro</returns>
    private int VerificaLinha(string placa)
    {
        //Declara váriavel do tipo int iniciando em 1, no excel a linha começa em 1
        int linha = 1;

        //Cria um objeto do tipo FileInfo e instancia passando no construtor da classe o caminho do arquivo excel
        FileInfo fi = new FileInfo(this.Arquivo);

        //Verifica se o arquivo existe no caminho informado
        if (!fi.Exists)
        {
            //Caso arquivo não exista cria um novo Woorbook
            ex.Workbooks.Add();
            //Retorna a linha 1 para a váriavel que chamou o metodo
            return linha;
        }

        //Caso arquivo exista abre o mesmo    
        ex.Workbooks.Open(this.Arquivo);

        //Percorre as linhas do excel, enquanto a linha não for nula(Em branco) continua
        while (ex.Cells[linha, 1].Value != null)
        {
            //Verifica se na linha informada e coluna 4 que é a coluna do cpf já existe algum cpf preenchido com o valor passado no parametro
            if (ex.Cells[linha, 5].Value.ToString() == placa)
            {
                //Caso encontre uma linha com a mesma placa retorna o valor da linha que a placa esta, 
                //desta forma é possível atualizar os dados do carro ao inves de criar um resgistro duplicado
                return linha;
            }
            //incrementa mais 1 a linha
            linha += 1;
        };

        //Caso não encontre a placa retorna a linha
        return linha;
    }

    public string[] ListaCarros(bool vendido){
        string[] carros = null;

        //Declara váriavel do tipo int iniciando em 1, no excel a linha começa em 1
        int linha = 1, linhaArray =0;

        //Cria um objeto do tipo FileInfo e instancia passando no construtor da classe o caminho do arquivo excel
        FileInfo fi = new FileInfo(this.Arquivo);

        //Verifica se o arquivo existe no caminho informado
        if (!fi.Exists)
        {
            //Caso arquivo não exista cria um novo Woorbook
            ex.Workbooks.Add();
            //Retorna a linha 1 para a váriavel que chamou o metodo
            return carros;
        }

        //Caso arquivo exista abre o mesmo    
        ex.Workbooks.Open(this.Arquivo);

        //Percorre as linhas do excel, enquanto a linha não for nula(Em branco) continua
        while (ex.Cells[linha, 1].Value != null)
        {
            if(!vendido){
                if (ex.Cells[linha, 6].Value.ToString() == "A venda")
                {
                    string dadoscarro = "";
                    for (int i = 0; i < 7; i++)
                    {
                        dadoscarro += ex.Cells[linha, i].Value.ToString() + ";";
                    }
                    carros[linhaArray] = dadoscarro;
                }
            }
            linha += 1;
            linhaArray +=1;
        };

        return carros;
    }
}