using System;

namespace sistema_concessionaria
{
    class Program
    {

        static string ArquivoCliente = @"C:\Users\fernando.henrique\Documents\aulas_backend\semana4\sistema_concessionaria\Clientes.xlsx";
        static string ArquivoCarro = @"C:\Users\fernando.henrique\Documents\aulas_backend\semana4\sistema_concessionaria\Carros.xlsx";

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
                        //Chama metodo para cadastrar novo cliente                        
                        CadastrarCliente();
                        break;
                    case 2:
                        CadastrarCarro();
                        break;
                    case 3:
                        RealizarVenda();
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

        static void CadastrarCliente(){
            //Cria um objeto do Tipo Cliente
            Cliente cliente = new Cliente(ArquivoCliente);
            cliente.ObterDados();
            cliente.Salvar();
        }

        static void CadastrarCarro(){
            //Cria um objeto do Tipo Carro
            Carro carro = new Carro(ArquivoCarro);
            carro.ObterDados();
            carro.Salvar();
        }

        static void RealizarVenda(){
            //Cria um objeto do Tipo Cliente
            Cliente cliente = new Cliente(ArquivoCliente);
            Carro carro = new Carro(ArquivoCarro);
            string[] dadosCliente;
            
            Console.WriteLine("Informe o cpf do cliente");
            string cpf = Console.ReadLine();

            //Carrega os dados do cliente caso exista
            dadosCliente = cliente.ObterDados(cpf);

            //Verifica se o cliente existe
            if(dadosCliente == null){
                //Caso cliente não existe informa para o usuário
                Console.WriteLine("Cpf não encontrado");
                //Chama o metodo para cadastrar novo cliente
                cliente.ObterDados();
                //Salva o novo cliente
                cliente.Salvar();
                //Carrega os dados do novo cliente
                dadosCliente = cliente.ObterDados(cpf);
            }

            //Retorna os carros não vendidos
            string[] carros = carro.ListaCarros(false);

        }
    }
}
