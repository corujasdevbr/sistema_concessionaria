using System;
using Validacao;
using NetOffice.ExcelApi;
using System.IO;

public class Cliente{

    Application ex = new Application();
    private string Nome { get; set; }
    private string Idade { get; set; }
    private string Email { get; set; }
    private string Cpf { get; set; }
    private Endereco Endereco { get; set; }

    public Cliente()
    {
         ex = new Application();
        Endereco = new Endereco();
    }

    public void ObterDados(){
        Documento doc = new Documento();

        Console.WriteLine("Informe o nome do cliente: ");
        this.Nome = Console.ReadLine();
        
        Console.WriteLine("Informe a idade do cliente: ");
        this.Idade = Console.ReadLine();
        
        Console.WriteLine("Informe o email do cliente: ");
        this.Email = Console.ReadLine();
        
        bool cpfValido = false;
        do{

            Console.WriteLine("Informe o cpf do cliente: ");
            this.Cpf = Console.ReadLine();  

            cpfValido = doc.ValidarCPF(this.Cpf);

            if(!cpfValido){
                Console.WriteLine("Cpf Inválido");
            }
        }while(!cpfValido);

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

    public void Salvar(string Arquivo){
        
        ex.DisplayAlerts = false;
        
        int linha = VerificaLinha(Arquivo, this.Cpf);

        ex.Cells[linha,1].Value = this.Nome;
        ex.Cells[linha,2].Value = this.Idade;
        ex.Cells[linha,3].Value = this.Email;
        ex.Cells[linha,4].Value = this.Cpf ;
        ex.Cells[linha,5].Value = this.Endereco.Logradouro;
        ex.Cells[linha,6].Value = this.Endereco.Numero;
        ex.Cells[linha,7].Value =  this.Endereco.Cep;
        ex.Cells[linha,8].Value =  this.Endereco.Complemento;
        ex.Cells[linha,9].Value =  this.Endereco.Bairro;
         
        ex.ActiveWorkbook.SaveAs(Arquivo);
        ex.Quit();
        ex.Dispose();
    }

    private int VerificaLinha(string Arquivo, string cpf){
        int linha = 1;

        FileInfo fi = new FileInfo(Arquivo);
        
        if(!fi.Exists){
            ex.Workbooks.Add();
            return linha;
        }
            
        ex.Workbooks.Open(Arquivo);
            
            while(ex.Cells[linha,1].Value != null){
                if(ex.Cells[linha,4].Value.ToString() == cpf){
                    break;
                }
                linha+=1;
            };

        return linha;
    }
}