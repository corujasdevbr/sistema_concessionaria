using System;
using Validacao;

public class Cliente{

    private string Nome { get; set; }
    private string Idade { get; set; }
    private string Email { get; set; }
    private string Cpf { get; set; }
    private Endereco Endereco { get; set; }

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

    
}