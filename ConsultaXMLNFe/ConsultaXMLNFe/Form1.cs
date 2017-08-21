using System;
using System.Net;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.Xml;
using System.Text;
using System.Windows.Forms;
using Recepcao = ConsultaXMLNFe.NfeRecepcao;
using Download = ConsultaXMLNFe.NfeDownload;
using Distribuicao = ConsultaXMLNFe.NfeDistribuicaoDFe;
using System.IO.Compression;
using System.Xml.Linq;

namespace ConsultaXMLNFe
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void CienciaOperacao(object sender, EventArgs e)
        {
            //
            //Ciencia da Operacao
            //
            var NFe_Rec = new Recepcao.RecepcaoEventoSoapClient();
            NFe_Rec.ClientCredentials.ClientCertificate.SetCertificate(StoreLocation.CurrentUser, StoreName.My, X509FindType.FindBySerialNumber, "‎SERIAL DO CERTIFICADO"); // <- sem espaços e tudo em caixa alta

            var notas = new string[] {"CHAVES DAS NOTAS"}; // este array não deve passar de 20 elementos, máximo permitido por lote de manifestação

            var sbXml = new StringBuilder();
            sbXml.Append(@"<?xml version=""1.0"" encoding=""UTF-8""?>
            <envEvento xmlns=""http://www.portalfiscal.inf.br/nfe"" versao=""1.00"">
            <idLote>1</idLote>
            ");
            foreach (var nota in notas)
            {
                sbXml.Append(@"
            <evento xmlns=""http://www.portalfiscal.inf.br/nfe"" versao=""1.00"">
            <infEvento Id=""ID210210" + nota + @"01"">
            <cOrgao>91</cOrgao>
            <tpAmb>1</tpAmb>
            <CNPJ>CNPJ DA EMPRESA</CNPJ>
            <chNFe>" + nota + @"</chNFe>
            <dhEvento>" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-ddTHH:mm:ss") + @"-03:00</dhEvento>
            <tpEvento>210210</tpEvento>
            <nSeqEvento>1</nSeqEvento>
            <verEvento>1.00</verEvento>
            <detEvento versao=""1.00"">
            <descEvento>Ciencia da Operacao</descEvento>
            </detEvento>
            </infEvento>
            </evento>
            ");
            }
            sbXml.Append("</envEvento>");

            var xml = new XmlDocument();
            xml.LoadXml(sbXml.ToString());

            var i = 0;
            foreach (var nota in notas)
            {
                var docXML = new SignedXml(xml);
                docXML.SigningKey = NFe_Rec.ClientCredentials.ClientCertificate.Certificate.PrivateKey;
                var refer = new Reference();
                refer.Uri = "#ID210210" + nota + "01";
                refer.AddTransform(new XmlDsigEnvelopedSignatureTransform());
                refer.AddTransform(new XmlDsigC14NTransform());
                docXML.AddReference(refer);

                var ki = new KeyInfo();
                ki.AddClause(new KeyInfoX509Data(NFe_Rec.ClientCredentials.ClientCertificate.Certificate));
                docXML.KeyInfo = ki;

                docXML.ComputeSignature();
                i++;
                xml.ChildNodes[1].ChildNodes[i].AppendChild(xml.ImportNode(docXML.GetXml(), true));
            }

            var NFe_Cab = new Recepcao.nfeCabecMsg();
            NFe_Cab.cUF = "35"; //SP => De acordo com a Tabela de Código de UF do IBGE
            NFe_Cab.versaoDados = "1.00";
            var resp = NFe_Rec.nfeRecepcaoEvento(NFe_Cab, xml);

            var fileResp = "D:\\Nfe\\" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + "-tempResp.xml";
            var fileReq = "D:\\Nfe\\" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + "-tempRequ.xml";
            File.WriteAllText(fileReq, xml.OuterXml);
            File.WriteAllText(fileResp, resp.OuterXml);
            System.Diagnostics.Process.Start(fileReq);
            System.Diagnostics.Process.Start(fileResp);
        }

        private void DownloadXML(object sender, EventArgs e)
        {
            //
            //Download da NFe
            // Para Download da xml por chave é necessario que a versão seja 1.01
            //
            var texto = @"<distDFeInt xmlns=""http://www.portalfiscal.inf.br/nfe"" versao = ""1.01"">
            <tpAmb>1</tpAmb>
            <cUFAutor>35</cUFAutor>
            <CNPJ>CNPJ DA EMPRESA</CNPJ>
            <consChNFe>
                <chNFe>CHAVE NOTA FISCAL</chNFe>
            </consChNFe>
            </distDFeInt>";

            // Pesquisar por NSU
            //var texto = @"<distDFeInt xmlns=""http://www.portalfiscal.inf.br/nfe"" versao = ""1.01"">
            //<tpAmb>1</tpAmb>
            //<cUFAutor>35</cUFAutor>
            //<CNPJ>CNPJ DA EMPRESA</CNPJ>
            //<distNSU>
            //    <ultNSU>000000000000000</ultNSU>
            //</distNSU>
            //</distDFeInt>";

            var xml = ConverterStringToXml(texto);
            BaixarXml(xml);
        }

        private void BaixarXml(XmlDocument xml)
        {
            var NFe_Sc = new Distribuicao.NFeDistribuicaoDFeSoapClient();
            NFe_Sc.ClientCredentials.ClientCertificate.SetCertificate(StoreLocation.CurrentUser, StoreName.My, X509FindType.FindBySerialNumber, "‎SERIAL DO CERTIFICADO");

            XElement x = XElement.Parse(xml.InnerXml);

            var arquivo = NFe_Sc.nfeDistDFeInteresse(x).ToString();

            var xmlNota = ConverterStringToXml(arquivo);
            var conteuZip = xmlNota.GetElementsByTagName("docZip")[0].InnerText;

            byte[] dados = Convert.FromBase64String(conteuZip);
            var xmlRetorno = descompactar(dados);
        }

        private XmlDocument ConverterStringToXml(string texto)
        {
            var sbXml = new StringBuilder();
            sbXml.Append(texto);
            var xml = new XmlDocument();
            xml.LoadXml(sbXml.ToString());
            return xml;
        }

        public string descompactar(byte[] conteudo)
        {
            GZipStream compressionStream = new GZipStream(new MemoryStream(conteudo), CompressionMode.Decompress);
            StreamReader reader = new StreamReader(compressionStream);
            string xml = "";
            foreach (var line in reader.ReadLine())
            {
                xml += line.ToString();
            }
            return xml;
        }
    }

}