using System;

namespace sistema_concessionaria
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                int opcao = 0;
            
                do{
                //Mostra um menu de opções para o usuário
                Console.WriteLine("Digite a opção");
                Console.WriteLine("1 - Cadastrar Cliente");
                Console.WriteLine("2 - Cadastrar Carro");
                Console.WriteLine("3 - Realizar Venda");
                Console.WriteLine("4 - Listar Carros Vendidos");
                Console.WriteLine("9 - Sair");

                //Recebe a opção do usuário
                opcao =  Int16.Parse(Console.ReadLine());
                
                //Verifica qual opção o usuário informou
                switch(opcao){
                    case 1:
                        //CadastrarCliente();
                        break;
                    case 2:
                        //CadastrarProduto();
                        break;
                    case 3:
                        //RealizarVenda();
                        break;
                    case 4:
                        //ExtratoCliente();
                        break;
                    case 9:{
                        //Pergunta para o usuário se ele realmente deseja sair
                        Console.WriteLine("Deseja realmente sair(s ou n)");
                        //Obtem a opção do usuário
                        string sair = Console.ReadLine();
                        //Verifica se ele digitou s
                        if(sair.ToLower().Contains("s"))
                            Environment.Exit(0);
                        else if(!sair.ToLower().Contains("n"))
                        {
                            opcao = 0;
                            Console.WriteLine("Opção Inválida");
                        }
                        else{
                            opcao = 0;
                        }
                        break;
                    }
                    default:
                        Console.WriteLine("Opção Inválida");
                        break;
                }
                //fica no laço até o usuário digitar 9
                }while(opcao != 9);
            }
            catch (System.Exception e)
            {
                //Caso ocorra algum erro grava no arquivo de erros
                Log log = new Log("Main", e.Message);
            }
        }
    }
}
