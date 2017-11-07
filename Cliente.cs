using System;
using Validacao;
using NetOffice.ExcelApi;
using System.IO;

public class Cliente
{

    //Declara uma váriavel do tipo Application para ser usada na classe
    Application ex;
    private string Nome { get; set; }
    private string Idade { get; set; }
    private string Email { get; set; }
    private string Cpf { get; set; }
    private Endereco Endereco { get; set; }
    private string Arquivo { get; set; }

    /// <summary>
    ///Construtor da classe com parametros 
    /// </summary>
    /// <param name="Arquivo"></param>
    public Cliente(string Arquivo)
    {
        //Define o caminho do arquivo excel dos clientes
        this.Arquivo = Arquivo;
        //Instancia a classe Application
        ex = new Application();
        //instancia a classe Endereço
        this.Endereco = new Endereco();
    }

    /// <summary>
    /// Metodo que obtem os dados do cliente
    /// </summary>
    public void ObterDados()
    {
        //Cria um objeto do tipo Documento para validação de cpf\cnpj
        Documento doc = new Documento();

        Console.WriteLine("Informe o nome do cliente: ");
        this.Nome = Console.ReadLine();

        Console.WriteLine("Informe a idade do cliente: ");
        this.Idade = Console.ReadLine();

        Console.WriteLine("Informe o email do cliente: ");
        this.Email = Console.ReadLine();

        //Declaração de variável do tipo bool para definir se sai ou continua no laço de repetição para validação de cpf\cnpj
        bool cpfValido = false;

        //Pede o cpf para o usuário
        do
        {

            Console.WriteLine("Informe o cpf do cliente: ");
            this.Cpf = Console.ReadLine();

            //Valida o cpf do usuário
            cpfValido = doc.ValidarCPF(this.Cpf);

            //Caso cpf seja inválido informa ao usuário e pede novamente
            if (!cpfValido)
            {
                Console.WriteLine("Cpf Inválido");
            }
            //Sai do laço apenas quando o cpf for válido
        } while (!cpfValido); //faça enquanto cpf == false

        Console.WriteLine("Informe o Endereço do cliente: ");
        this.Endereco.Logradouro = Console.ReadLine();

        Console.WriteLine("Informe o Número do cliente: ");
        this.Endereco.Numero = Console.ReadLine();

        Console.WriteLine("Informe o Cep do cliente: ");
        this.Endereco.Cep = Console.ReadLine();

        Console.WriteLine("Informe o Complemento do cliente: ");
        this.Endereco.Complemento = Console.ReadLine();

        Console.WriteLine("Informe o Bairro do cliente: ");
        this.Endereco.Bairro = Console.ReadLine();
    }

    /// <summary>
    /// Obtem os dados do cliente passado no parametro
    /// </summary>
    /// <param name="cpf">Parametro do tipo string com os dados do cliente</param>
    /// <returns>Retorna um array com os dados do cliente</returns>
    public string[] ObterDados(string cpf)
    {
        //Declara váriavel do tipo int iniciando em 1, no excel a linha começa em 1
        int linha = 1;

        //Cria um objeto do tipo FileInfo e instancia passando no construtor da classe o caminho do arquivo excel
        FileInfo fi = new FileInfo(this.Arquivo);

        //Verifica se o arquivo existe no caminho informado
        if (!fi.Exists)
        {
            //Retorna null pois o arquivo não existe e não possui nenhum cliente
            return null;
        }

        //Caso arquivo exista abre o mesmo    
        ex.Workbooks.Open(this.Arquivo);

        //Percorre as linhas do excel, enquanto a linha não for nula(Em branco) continua
        while (ex.Cells[linha, 1].Value != null)
        {
            //Verifica se na linha informada e coluna 4 que é a coluna do cpf já existe algum cpf preenchido com o valor passado no parametro
            if (ex.Cells[linha, 4].Value.ToString() == cpf)
            {
                //Caso encontre uma linha com o mesmo cpf cria um array para armazenar os dados do cliente, 
                
                string[] dadosCliente = new string[9];
                //Carrega o array com os dados do cliente
                dadosCliente[0] = ex.Cells[linha, 0].Value.ToString();
                dadosCliente[1] = ex.Cells[linha, 1].Value.ToString();
                dadosCliente[2] = ex.Cells[linha, 2].Value.ToString();
                dadosCliente[3] = ex.Cells[linha, 3].Value.ToString();
                dadosCliente[4] = ex.Cells[linha, 4].Value.ToString();
                dadosCliente[5] = ex.Cells[linha, 5].Value.ToString();
                dadosCliente[6] = ex.Cells[linha, 6].Value.ToString();
                dadosCliente[7] = ex.Cells[linha, 7].Value.ToString();
                dadosCliente[8] = ex.Cells[linha, 8].Value.ToString();
                
                //retorna para o metodo que chamou os dados do cliente
                return dadosCliente;
            }
            //incrementa mai 1 a linha
            linha += 1;
        };

        //Caso não encontre o cpf retorna nulo
        return null;
    }

    /// <summary>
    /// Salva os dados obtidos no metodo ObterDados
    /// </summary>
    public void Salvar()
    {
        //Não mostra nenhuma modal
        ex.DisplayAlerts = false;

        //verifica em qual linha deve ser salvo o cliente, cliente novo salva um novo e cliente já cadastrado altera os dados
        int linha = VerificaLinha(this.Cpf);

        //Prenche os dados na linha obtida acima e nas colunas definidas
        ex.Cells[linha, 1].Value = this.Nome;
        ex.Cells[linha, 2].Value = this.Idade;
        ex.Cells[linha, 3].Value = this.Email;
        ex.Cells[linha, 4].Value = this.Cpf;
        ex.Cells[linha, 5].Value = this.Endereco.Logradouro;
        ex.Cells[linha, 6].Value = this.Endereco.Numero;
        ex.Cells[linha, 7].Value = this.Endereco.Cep;
        ex.Cells[linha, 8].Value = this.Endereco.Complemento;
        ex.Cells[linha, 9].Value = this.Endereco.Bairro;

        //Salva o arquivo no local informado
        ex.ActiveWorkbook.SaveAs(this.Arquivo);
        //fecha o Application\Excel
        ex.Quit();
        //Retira da memoria
        ex.Dispose();
    }

    /// <summary>
    /// Verifica qual linha irá incluir\alterar o cliente
    /// </summary>
    /// <param name="cpf">CPF do cliente, será único</param>
    /// <returns></returns>
    private int VerificaLinha(string cpf)
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
            if (ex.Cells[linha, 4].Value.ToString() == cpf)
            {
                //Caso encontre uma linha com o mesmo cpf retorna o valor da linha que o cpf esta, 
                //desta forma é possível atualiar os dados do cliente ao inves de criar um resgistro duplicado
                return linha;
            }
            //incrementa mai 1 a linha
            linha += 1;
        };

        //Caso não encontre o cpf retorna a linha
        return linha;
    }
}