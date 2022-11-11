using Newtonsoft.Json;
using RestSharp;
using RestSharp.Authenticators;
using System.Collections.Generic;

using IronXL;
using System.Text.Json;

namespace Teste
{
    class Program
    {
        private const string token = ""; // inserir a chave enviado por email
        private const string Empresa = "https://crm.rdstation.com/api/v1/organizations?token=" + token;
        private const string Cliente = "https://crm.rdstation.com/api/v1/contacts?token=" + token;

        static void Main(string[] args)
        {
            #region Conexao Empresa
            var clientEmpresa = new RestClient(Empresa);
            clientEmpresa.Authenticator = new HttpBasicAuthenticator("alfa-erp", "alfa.123");
            var requestEmpresa = new RestRequest(Empresa, Method.Get);
            requestEmpresa.AddHeader("Cookie", "_session_id=7df80d0143fbdb04e851af3e7e25eb95");
            RestResponse responseEmpresa = clientEmpresa.Execute(requestEmpresa);
            var result = JsonConvert.DeserializeObject<dynamic>(responseEmpresa.Content);
            List<Empresa> organization = result.organizations.ToObject<List<Empresa>>();
            #endregion

            #region Conexao Cliente
            var clientContato = new RestClient(Cliente);
            clientContato.Authenticator = new HttpBasicAuthenticator("alfa-erp", "alfa.123");
            var requestContato = new RestRequest(Cliente, Method.Get);
            requestContato.AddHeader("Cookie", "_session_id=7df80d0143fbdb04e851af3e7e25eb95");
            RestResponse responseContato = clientEmpresa.Execute(requestEmpresa);
            var resultContato = JsonConvert.DeserializeObject<dynamic>(responseEmpresa.Content);
            List<Contatos> listaContatos = result.organizations.ToObject<List<Contatos>>();            
            #endregion

            #region Excel            
            WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
            var sheet = workbook.CreateWorkSheet("teste");

            sheet["A1"].Value = "CodigoCliente";
            sheet["B1"].Value = "Empresa";
            sheet["C1"].Value = "Segmento";
            sheet["D1"].Value = "Url";
            sheet["E1"].Value = "Resumo";
            sheet["F1"].Value = "Contato";
            sheet["G1"].Value = "Cargo";
            sheet["H1"].Value = "Telefone";
            sheet["I1"].Value = "Email1";
            sheet["J1"].Value = "Email2";

            int i = 0;
            foreach (var empresa in organization)
            {
                sheet["B2"].Value = empresa.NomeEmpresa;
                sheet["D2"].Value = empresa.Url;
                sheet["E2"].Value = empresa.Resumo;

                foreach (var segmento in empresa.EmpresaSegmento)
                {
                    sheet["C2"].Value = segmento.Segmento;
                }
                foreach (var contato in empresa.Contatos)
                {
                    
                    sheet["F2"].Value = contato.Contato;
                    sheet["G2"].Value = contato.Cargo;
                    foreach (var telefone in contato.Telefones)
                    {
                        sheet["H2"].Value = "+" + telefone.Telefone;

                    }
                    foreach (var email in contato.Emails)
                    {
                        if (i == 0)
                        {
                            sheet["I2"].Value = email.Email;
                        }
                        if (i == 1)
                        {
                            sheet["J2"].Value = email.Email;
                        }
                        i++;
                    }
                }
            }
            foreach (var dadosContato in listaContatos)
            {
                sheet["A2"].Value = dadosContato.CodigoCliente;
            }
            workbook.SaveAs($@"C:\Temp\teste.xlsx");
            #endregion
        }
    }
    public class Empresa
    {
        [JsonProperty("name")]
        public string NomeEmpresa { get; set; }
        [JsonProperty("resume")]
        public string Resumo { get; set; }
        [JsonProperty("url")]
        public string Url { get; set; }
        [JsonProperty("contacts")]
        public List<Contatos> Contatos { get; set; }
        [JsonProperty("organization_segments")]
        public List<EmpresaSegmento> EmpresaSegmento { get; set; }
    }
    public class EmpresaSegmento
    {
        [JsonProperty("name")]
        public string Segmento { get; set; }
    }
    public class Contatos
    {
        [JsonProperty("id")]
        public string CodigoCliente { get; set; }
        [JsonProperty("name")]
        public string Contato { get; set; }
        [JsonProperty("title")]
        public string Cargo { get; set; }
        [JsonProperty("organization_id")]
        public string IdEmpresa { get; set; }
        [JsonProperty("emails")]
        public List<Emails> Emails { get; set; }
        [JsonProperty("phones")]
        public List<Telefones> Telefones { get; set; }
    }
    public class Emails
    {
        [JsonProperty("email")]
        public string Email { get; set; }
    }
    public class Telefones
    {
        [JsonProperty("phone")]
        public string Telefone { get; set; }
    }
}
