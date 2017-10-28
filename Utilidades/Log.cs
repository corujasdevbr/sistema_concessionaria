using System.IO;
/// <summary>
/// Classe log para gravar os erros da aplicação
/// </summary>
public class Log{

    private string Metodo { get; set; }     
    private string Mensagem { get; set; }     

    public Log()
    {
        
    }

    /// <summary>
    /// Construtor que recebe parametros de entrada e já salva no arquivo
    /// </summary>
    /// <param name="Metodo">Metodo que ocorreu o erro</param>
    /// <param name="Mensagem">Mensagem de erro</param>
    public Log(string Metodo, string Mensagem)
    {
        this.Metodo = Metodo;
        this.Mensagem = Mensagem;
        SalvarLog();
    }

    private void SalvarLog(){
        StreamWriter sw = new StreamWriter("logerro.txt");
        sw.WriteLine("Metodo: " + this.Metodo + " - erro: " + this.Mensagem);
        sw.Close();
    }
}