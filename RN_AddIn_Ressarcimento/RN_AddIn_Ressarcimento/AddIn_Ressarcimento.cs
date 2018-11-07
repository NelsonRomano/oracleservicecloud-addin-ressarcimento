using RightNow.AddIns.AddInViews;
using System;
using System.AddIn;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using WebServiceCall;

namespace RN_AddIn_Ressarcimento {
    public class WorkspaceRibbonAddIn : Panel, IWorkspaceRibbonButton {

        private IRecordContext RecordContext {
            get; set;
        }

        public WorkspaceRibbonAddIn(bool inDesignMode, IRecordContext RecordContext) {
            this.RecordContext = RecordContext;
        }

        #region IWorkspaceRibbonButton Members
        public new void Click() {
            gravarArquivo();
        }
        #endregion

        public static string extrairValor(string xmlResposta) {
            string valor = "";
            XmlDocument document = new XmlDocument();
            document.LoadXml(xmlResposta);
            var namespaceManager = new XmlNamespaceManager(document.NameTable);
            namespaceManager.AddNamespace("soapenv", "http://schemas.xmlsoap.org/soap/envelope/");
            namespaceManager.AddNamespace("n0", "urn:messages.ws.rightnow.com/v1_3");
            namespaceManager.AddNamespace("xsi", "http://www.w3.org/2001/XMLSchema-instance");
            valor = document.SelectSingleNode("//n0:Row", namespaceManager).InnerText;
            valor = Regex.Replace(valor, "^\"(.+?)\"$", "$1");
            if(valor.Equals("")) {
                return "N/I";
            }
            else {
                return valor;
            }
        }

        public void gravarArquivo() {
            IIncident inc = (IIncident) RecordContext.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Incident);
            IContact contato = (IContact) RecordContext.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Contact);
            string incidentID = inc.ID.ToString();
            string contatoID = contato.ID.ToString();
            string dataAtualizacao = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Incident.UpdatedTime FROM Incident WHERE Incident.Id = " + incidentID));
            string produto = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.CO.Produto.Name FROM Incident WHERE Incident.id = " + incidentID));
            string categoria = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Incident.Category.Name FROM Incident WHERE Incident.Id = " + incidentID));
            string quantidade = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.quantidade FROM Incident WHERE Incident.id = " + incidentID));
            string valor = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.rs FROM Incident WHERE Incident.id = " + incidentID));
            string vencimento = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.vencimento FROM Incident WHERE Incident.id = " + incidentID));
            string idOrganizacao = pegarIDOrganizacao(incidentID);
            string labelCpfCnpj = "";

            string nome;
            string cpfCnpj;
            string rua;
            string cep;
            string cidade;
            string estado;
            string telefone;
            string celular;

            if(idOrganizacao != null) {
                nome = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Organization.Name FROM Organization WHERE Organization.Id = " + idOrganizacao));
                cpfCnpj = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.cnpj FROM Organization WHERE Organization.Id = " + idOrganizacao));
                labelCpfCnpj = "CNPJ";

                if(cpfCnpj != null && !cpfCnpj.Equals("N/I") && !cpfCnpj.Equals("") && cpfCnpj.Length == 14) {                    
                    cpfCnpj = cpfCnpj.Substring(0, 2) + "." + cpfCnpj.Substring(2, 3) + "." + cpfCnpj.Substring(5, 3) + "/" + cpfCnpj.Substring(8, 4) + "-" + cpfCnpj.Substring(12, 2);
                }

                rua = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Addresses.street FROM Organization WHERE Organization.Id = " + idOrganizacao));

                if(rua != null && !rua.Equals("") && !rua.Equals("N/I")) {
                    rua = rua.Replace("\"", "");
                }

                cep = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Addresses.PostalCode FROM Organization WHERE Organization.Id = " + idOrganizacao));
                cidade = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Addresses.City FROM Organization WHERE Organization.Id = " + idOrganizacao));
                estado = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Addresses.StateOrProvince.Name FROM Organization WHERE Organization.Id = " + idOrganizacao));
                telefone = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.telefone_1 FROM Organization WHERE Organization.Id = " + idOrganizacao));
                celular = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.telefone_2 FROM Organization WHERE Organization.Id = " + idOrganizacao));
            }

            else {
                nome = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Contact.Name.First FROM Contact WHERE Contact.id = " + contatoID));
                cpfCnpj = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.cpf FROM Contact WHERE Contact.id = " + contatoID));
                labelCpfCnpj = "CPF";

                if(cpfCnpj != null && !cpfCnpj.Equals("") && !cpfCnpj.Equals("N/I") && cpfCnpj.Length == 11) {                    
                    cpfCnpj = cpfCnpj.Substring(0, 3) + "." + cpfCnpj.Substring(3, 3) + "." + cpfCnpj.Substring(6, 3) + "-" + cpfCnpj.Substring(9, 2);
                }

                rua = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Contact.Address.Street FROM Contact WHERE Contact.id = " + contatoID));

                if(rua != null && !rua.Equals("") && !rua.Equals("N/I")) {
                    rua = rua.Replace("\"", "");
                }

                cep = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Contact.Address.PostalCode FROM Contact WHERE Contact.id = " + contatoID));
                cidade = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Contact.Address.City FROM Contact WHERE Contact.id = " + contatoID));
                estado = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Contact.Address.StateOrProvince.Name FROM Contact WHERE Contact.id = " + contatoID));
                telefone = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Contact.Phones.Number FROM Contact WHERE Contact.Phones.PhoneType = 4 AND Contact.Id = " + contatoID));
                celular = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.telefone_celular FROM Contact WHERE Contact.Id = " + contatoID));
            }

            string banco = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.banco FROM Contact WHERE Contact.id = " + contatoID));
            string agencia = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.agencia FROM Contact WHERE Contact.id = " + contatoID));
            string contaCorrente = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.conta_corrente FROM Contact WHERE Contact.id = " + contatoID));

            if(telefone == null || telefone.Equals("N/I") || telefone.Equals("")) {
                telefone = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Contact.Phones.Number FROM Contact WHERE Contact.Phones.PhoneType = 3 AND Contact.Id = " + contatoID));
            }

            if(telefone != null && !telefone.Equals("N/I") && !telefone.Equals("")) {
                if(telefone.Length == 10) {
                    telefone = telefone.Substring(0, 2) + " " + telefone.Substring(2, 4) + "-" + telefone.Substring(6, 4);
                }

                if(telefone.Length == 11) {
                    telefone = telefone.Substring(0, 2) + " " + telefone.Substring(2, 5) + "-" + telefone.Substring(7, 4);
                }
            }

            if(celular != null && !celular.Equals("N/I") && !celular.Equals("")) {
                if(celular.Length == 10) {
                    celular = celular.Substring(0, 2) + " " + celular.Substring(2, 4) + "-" + celular.Substring(6, 4);
                }

                if(celular.Length == 11) {
                    celular = celular.Substring(0, 2) + " " + celular.Substring(2, 5) + "-" + celular.Substring(7, 4);
                }
            }

            if(quantidade != null && !quantidade.Equals("N/I") && !quantidade.Equals("")) {
                if(Convert.ToInt32(quantidade) > 1) {
                    quantidade = quantidade + " unidades";
                }
                else {
                    quantidade = quantidade + " unidade";
                }
            }

            if(valor != null && !valor.Equals("N/I") && !valor.Equals("")) {
                valor = "R$ " + valor;
            }

            DateTime data;
            if(!DateTime.TryParse(dataAtualizacao, out data)) {
                MessageBox.Show("O campo Data de Criação não foi preenchido com um valor válido.");
            }
            else {
                dataAtualizacao = data.ToString("dd/MM/yyyy HH:mm:ss");
            }

            if(cep != null && !cep.Equals("N/I") && !cep.Equals("") && cep.Length == 8) {
                cep = cep.Substring(0, 5) + "-" + cep.Substring(5, 3);
            }

            string html1 = "<!DOCTYPE html>"
                + "<html xmlns=\"undefined\">"
                + "  <head>"
                + "    <meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\">"
                + "      <title>Ressarcimento</title>"
                + "      <link rel=\"themeData\" href=\"file:///C:/Users/nromano/Desktop/Carta%20de%20ressarcimento_arquivos/themedata.thmx\">"
                + "        <link rel=\"colorSchemeMapping\" href=\"file:///C:/Users/nromano/Desktop/Carta%20de%20ressarcimento_arquivos/colorschememapping.xml\">"
                + "          <style type=\"text/css\" xml:space=\"preserve\"> @font-face { font-family: Cambria Math; } @font-face { font-family: Calibri; } @page WordSection1 { size: 595.3pt 841.9pt; mso-fareast-font-family: Calibri; mso-bidi-font-family: \"Times New Roman\"; mso-fareast-theme-font: minor-latin; mso-bidi-theme-font: minor-bidi; mso-fareast-language: EN-US; mso-ascii-font-family: Calibri; mso-hansi-font-family: Calibri; mso-style-type: export-only; mso-default-props: yes; mso-ascii-theme-font: minor-latin; mso-hansi-theme-font: minor-latin; } .MsoPapDefault { -BOTTOM: 8pt; LINE-HEIGHT: 107%; mso-style-type: export-only; } DIV.WordSection1 { page: WordSection1; } .style2 { font-size: 18px; font-family: Arial, Helvetica, sans-serif; } body { margin-left: 5px; } .style3 { font-family: Arial, Helvetica, sans-serif; } </style>"
                + "        </head>"
                + "        <body bgcolor=\"#FFFFFF\" style=\"tab-interval: 35.4pt\" lang=\"PT-BR\" xml:lang=\"PT-BR\">"
                + "          <div class=\"WordSection1\">"
                + "            <p class=\"MsoNormal\" style=\"MARGIN-BOTTOM: 5pt; MARGIN-RIGHT: 85pt;\" align=\"left\">"
                + "              <span style=\"FONT-FAMILY: &quot;Century Gothic&quot;,sans-serif; COLOR: #181717\">"
                + "                <img src=\"https://superarx.custhelp.com/euf/assets/images/Logo.png\" alt=\"Image\" width=\"232\" height=\"35\" style=\"width:214px;height:27px;\" border=\"0\">"
                + "                </span>"
                + "              </p>"
                + "              <table width=\"1112\" height=\"352%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" class=\"MsoNormalTable\" style=\"WIDTH: 647pt;/* BORDER-COLLAPSE: collapse; */mso-yfti-tbllook: 1184;mso-padding-alt: 0cm 3.5pt 0cm 3.5pt;\">"
                + "                <tbody>"
                + "                  <tr style=\"HEIGHT: 15pt; mso-yfti-irow: 0; mso-yfti-firstrow: yes\">"
                + "                    <td width=\"215\" height=\"2%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 15pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\"/>"
                + "                    <td width=\"303\" colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"HEIGHT: 15pt; WIDTH: 154pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td width=\"76\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"HEIGHT: 15pt; WIDTH: 47pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 113pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"215\" nowrap=\"nowrap\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 107pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"303\" nowrap=\"NOWRAP\">      </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 15pt; mso-yfti-irow: 1\">"
                + "                    <td width=\"215\" height=\"2%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 15pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">&nbsp;</td>"
                + "                    <td width=\"303\" colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"HEIGHT: 15pt; WIDTH: 154pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 47pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" nowrap=\"NOWRAP\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 113pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"215\" nowrap=\"nowrap\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 107pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"303\" nowrap=\"NOWRAP\">      </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 15.75pt; mso-yfti-irow: 2\">"
                + "                    <td height=\"2%\" colspan=\"6\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 15.75pt; WIDTH: 537pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">"
                + "                      <p align=\"center\" class=\"MsoNormal style2\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: center; LINE-HEIGHT: normal\">"
                + "                        <b>"
                + "                          <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Ressarcimento de produto com desvio qualidade - Depósito em conta</span>"
                + "                        </b>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 15pt; mso-yfti-irow: 3\">"
                + "                    <td width=\"215\" height=\"2%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 15pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td width=\"303\" colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"HEIGHT: 15pt; WIDTH: 154pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 47pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" nowrap=\"NOWRAP\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 113pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"215\" nowrap=\"nowrap\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 107pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"303\" nowrap=\"NOWRAP\">      </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 15pt; mso-yfti-irow: 4\">"
                + "                    <td width=\"215\" height=\"2%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 15pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">"
                + "                      <p align=\"right\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Data:</span>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td width=\"303\" colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"FONT-SIZE: 12pt; TEXT-DECORATION: none; HEIGHT: 15pt; FONT-FAMILY: &quot;Arial&quot;,sans-serif; WIDTH: 154pt; VERTICAL-ALIGN: auto; FONT-WEIGHT: 400; COLOR: #000000; PADDING-BOTTOM: 0cm; FONT-STYLE: normal; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt; BACKGROUND-COLOR: #ffffff\">"
                + "                      <span class=\"style3\">" + dataAtualizacao + "</span>"
                + "                    </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 47pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" nowrap=\"NOWRAP\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 113pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"215\" nowrap=\"nowrap\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 107pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"303\" nowrap=\"NOWRAP\">";

            string html2 = "                      <p class=\"MsoNormal\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\" align=\"right\">"
                + "                        <span class=\"GramE style3\">"
                + "                          <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">nº " + incidentID + "</span>"
                + "                        </span>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 15pt; mso-yfti-irow: 5\">"
                + "                    <td width=\"215\" height=\"2%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 15pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td width=\"303\" colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"HEIGHT: 15pt; WIDTH: 154pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 47pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" nowrap=\"NOWRAP\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 113pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"215\" nowrap=\"nowrap\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 107pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"303\" nowrap=\"NOWRAP\">      </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 15.75pt; mso-yfti-irow: 6\">"
                + "                    <td height=\"2%\" colspan=\"6\" valign=\"bottom\" nowrap=\"nowrap\" style=\"BORDER-TOP: medium none; HEIGHT: 15.75pt; BORDER-RIGHT: medium none; WIDTH: 537pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-bottom-alt: solid windowtext .5pt\">"
                + "                      <p align=\"center\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: center; LINE-HEIGHT: normal\">"
                + "                        <b>"
                + "                          <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Dados do Produto:</span>"
                + "                        </b>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 5.25pt; mso-yfti-irow: 7\">"
                + "                    <td width=\"215\" height=\"1%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 5.25pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td width=\"303\" colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"HEIGHT: 5.25pt; WIDTH: 154pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td style=\"HEIGHT: 5.25pt; WIDTH: 47pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" nowrap=\"NOWRAP\">        </td>"
                + "                    <td style=\"HEIGHT: 5.25pt; WIDTH: 113pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"215\" nowrap=\"nowrap\">        </td>"
                + "                    <td style=\"HEIGHT: 5.25pt; WIDTH: 107pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"303\" nowrap=\"NOWRAP\">      </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 26.25pt; mso-yfti-irow: 8\">"
                + "                    <td width=\"215\" height=\"4%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 26.25pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">"
                + "                      <p align=\"right\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Produto:</span>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td width=\"303\" colspan=\"5\" valign=\"bottom\" nowrap=\"\" style=\"BORDER-TOP: medium none; HEIGHT: 26.25pt; BORDER-RIGHT: medium none; WIDTH: 154pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-bottom-alt: solid windowtext .5pt\">"
                + "                      <p align=\"left\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\">"
                + "                        <b>"
                + "                          <span style=\"margin-left: 5px; COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">" + produto + "</span>"
                + "                        </b>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 24.95pt; mso-yfti-irow: 9\">"
                + "                    <td width=\"215\" height=\"3%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 24.95pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">"
                + "                      <p align=\"right\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Motivo:</span>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td width=\"303\" colspan=\"5\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"BORDER-TOP: medium none; HEIGHT: 24.95pt; BORDER-RIGHT: medium none; WIDTH: 154pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-bottom-alt: solid windowtext .5pt\">"
                + "                      <p align=\"left\" class=\"MsoNormal style3\" style=\"margin-left: 5px; MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\">"
                + "                        <strong>" + categoria + "</strong>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 24.95pt; mso-yfti-irow: 10\">"
                + "                    <td width=\"215\" height=\"3%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 24.95pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">"
                + "                      <p align=\"right\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Quantidade:</span>"
                + "                      </p>";

            string html3 = "                    </td>"
                + "                    <td width=\"303\" colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"BORDER-TOP: medium none; HEIGHT: 24.95pt; BORDER-RIGHT: medium none; WIDTH: 154pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-bottom-alt: solid windowtext .5pt\">"
                + "                      <p align=\"left\" class=\"MsoNormal style3\" style=\"margin-left: 5px; MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\">"
                + "                        <strong>" + quantidade + "</strong>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 24.75pt; mso-yfti-irow: 11\">"
                + "                    <td width=\"215\" height=\"3%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 24.75pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">"
                + "                      <p class=\"MsoNormal\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\" align=\"right\">"
                + "                        <span class=\"GramE style3\">"
                + "                          <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Valor:</span>"
                + "                        </span>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td width=\"303\" colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"BORDER-TOP: medium none; HEIGHT: 24.75pt; BORDER-RIGHT: medium none; WIDTH: 154pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-bottom-alt: solid windowtext .5pt\">"
                + "                      <p align=\"left\" class=\"MsoNormal style3\" style=\"margin-left: 5px; MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\">"
                + "                        <strong>" + valor + "</strong>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 24.75pt; mso-yfti-irow: 11\">"
                + "                    <td width=\"215\" height=\"3%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 24.75pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">"
                + "                      <p class=\"MsoNormal\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\" align=\"right\">"
                + "                        <span class=\"GramE style3\">"
                + "                          <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Vencimento:</span>"
                + "                        </span>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td width=\"303\" colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"BORDER-TOP: medium none; HEIGHT: 24.75pt; BORDER-RIGHT: medium none; WIDTH: 154pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-bottom-alt: solid windowtext .5pt\">"
                + "                      <p align=\"left\" class=\"MsoNormal style3\" style=\"margin-left: 5px; MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\">"
                + "                        <strong>" + vencimento + "</strong>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 33pt; mso-yfti-irow: 13\">"
                + "                    <td height=\"4%\" colspan=\"6\" valign=\"bottom\" nowrap=\"nowrap\" style=\"BORDER-TOP: medium none; HEIGHT: 33pt; BORDER-RIGHT: medium none; WIDTH: 537pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-bottom-alt: solid windowtext .5pt\">"
                + "                      <p align=\"center\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: center; LINE-HEIGHT: normal\">"
                + "                        <b>"
                + "                          <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Dados do Cliente:</span>"
                + "                        </b>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 15.75pt; mso-yfti-irow: 14\">"
                + "                    <td width=\"215\" height=\"2%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 15.75pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td width=\"303\" colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"HEIGHT: 15.75pt; WIDTH: 154pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td style=\"HEIGHT: 15.75pt; WIDTH: 47pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" nowrap=\"NOWRAP\">        </td>"
                + "                    <td style=\"HEIGHT: 15.75pt; WIDTH: 113pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"215\" nowrap=\"nowrap\">        </td>"
                + "                    <td style=\"HEIGHT: 15.75pt; WIDTH: 107pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"303\" nowrap=\"NOWRAP\">      </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 15.75pt; mso-yfti-irow: 15\">"
                + "                    <td width=\"215\" height=\"2%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 15.75pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">"
                + "                      <p align=\"right\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">CLIENTE</span>:</p>"
                + "                    </td>"
                + "                    <td colspan=\"5\" valign=\"bottom\" nowrap=\"\" style=\"BORDER-TOP: medium none; HEIGHT: 15.75pt; BORDER-RIGHT: medium none; WIDTH: 421pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-bottom-alt: solid windowtext .5pt\">"
                + "                      <p align=\"left\" class=\"MsoNormal style3\" style=\"margin-left: 5px; MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\">"
                + "                        <strong>" + nome + "</strong>";

            string html4 = "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 24.95pt; mso-yfti-irow: 16\">"
                + "                    <td width=\"215\" height=\"3%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 24.95pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">"
                + "                      <p align=\"right\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">" + labelCpfCnpj + ":</span>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td width=\"303\" colspan=\"5\" valign=\"bottom\" nowrap=\"\" style=\"BORDER-TOP: medium none; HEIGHT: 24.95pt; BORDER-RIGHT: medium none; WIDTH: 154pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-bottom-alt: solid windowtext .5pt\">"
                + "                      <p align=\"left\" class=\"MsoNormal style3\" style=\"margin-left: 5px; MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\">"
                + "                        <strong>" + cpfCnpj + "</strong>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 24.95pt; mso-yfti-irow: 17\">"
                + "                    <td width=\"215\" height=\"3%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 24.95pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">"
                + "                      <p align=\"right\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Endereço:</span>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td colspan=\"10\" width=\"303\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"BORDER-TOP: medium none; HEIGHT: 24.95pt; BORDER-RIGHT: medium none; WIDTH: 154pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-bottom-alt: solid windowtext .5pt\">"
                + "                      <p align=\"left\" class=\"MsoNormal style3\" style=\"margin-left: 5px; MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\">"
                + "                        <strong>" + rua + "</strong>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 24.95pt; mso-yfti-irow: 18\">"
                + "                    <td width=\"215\" height=\"3%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 24.95pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">"
                + "                      <p align=\"right\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">CEP:</span>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td width=\"303\" colspan=\"5\" valign=\"bottom\" nowrap=\"\" style=\"BORDER-TOP: medium none; HEIGHT: 24.95pt; BORDER-RIGHT: medium none; WIDTH: 154pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-bottom-alt: solid windowtext .5pt\">"
                + "                      <p align=\"left\" class=\"MsoNormal style3\" style=\"margin-left: 5px; MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\">"
                + "                        <strong>" + cep + "</strong>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 24.95pt; mso-yfti-irow: 19\">"
                + "                    <td width=\"215\" height=\"3%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 24.95pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">"
                + "                      <p align=\"right\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Cidade / UF:</span>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td width=\"303\" colspan=\"5\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"BORDER-TOP: medium none; HEIGHT: 24.95pt; BORDER-RIGHT: medium none; WIDTH: 154pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-bottom-alt: solid windowtext .5pt\">"
                + "                      <p align=\"left\" class=\"MsoNormal style3\" style=\"margin-left: 5px; MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\">"
                + "                        <strong>" + cidade + " / " + estado + "</strong>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 24.95pt; mso-yfti-irow: 20\">"
                + "                    <td width=\"215\" height=\"3%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 24.95pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">"
                + "                      <p align=\"right\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Telefone:</span>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td width=\"303\" colspan=\"5\" valign=\"bottom\" nowrap=\"\" style=\"BORDER-TOP: medium none; HEIGHT: 24.95pt; BORDER-RIGHT: medium none; WIDTH: 154pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-bottom-alt: solid windowtext .5pt\">"
                + "                      <p align=\"left\" class=\"MsoNormal style3\" style=\"margin-left: 5px; MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\">"
                + "                        <strong>" + telefone + "</strong>";

            string html5 = "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 15.75pt; mso-yfti-irow: 15\">"
                + "                    <td width=\"215\" height=\"2%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 15.75pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">"
                + "                      <p align=\"right\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Celular</span>:</p>"
                + "                    </td>"
                + "                    <td colspan=\"5\" valign=\"bottom\" nowrap=\"\" style=\"BORDER-TOP: medium none; HEIGHT: 15.75pt; BORDER-RIGHT: medium none; WIDTH: 421pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-bottom-alt: solid windowtext .5pt\">"
                + "                      <p align=\"left\" class=\"MsoNormal style3\" style=\"margin-left: 5px; MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\">"
                + "                        <strong>" + celular + "</strong>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 24.95pt; mso-yfti-irow: 22\">"
                + "                    <td height=\"3%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 24.95pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"HEIGHT: 24.95pt; WIDTH: 154pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td style=\"HEIGHT: 24.95pt; WIDTH: 47pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" nowrap=\"NOWRAP\">        </td>"
                + "                    <td style=\"HEIGHT: 24.95pt; WIDTH: 113pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" nowrap=\"nowrap\">        </td>"
                + "                    <td style=\"HEIGHT: 24.95pt; WIDTH: 107pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" nowrap=\"NOWRAP\">      </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 15pt; mso-yfti-irow: 23\">"
                + "                    <td width=\"215\" height=\"32\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 15pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\"/>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 15pt; mso-yfti-irow: 24\">"
                + "                    <td width=\"215\" height=\"2%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"BORDER-TOP: windowtext 1pt solid;HEIGHT: 15pt;BORDER-RIGHT: windowtext 1pt solid;WIDTH: 116pt;BORDER-BOTTOM: medium none;PADDING-BOTTOM: 0cm;PADDING-TOP: 0cm;PADDING-LEFT: 3.5pt;BORDER-LEFT: windowtext 1pt solid;PADDING-RIGHT: 3.5pt;mso-border-top-alt: solid windowtext .5pt;mso-border-left-alt: solid windowtext .5pt;\" colspan=\"3\">"
                + "                      <p class=\"MsoNormal\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\" align=\"right\">&nbsp;</p>"
                + "                    </td>"
                + "                    <td width=\"215\" height=\"2%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"BORDER-TOP: windowtext 1pt solid;HEIGHT: 15pt;BORDER-RIGHT: windowtext 1pt solid;WIDTH: 116pt;BORDER-BOTTOM: medium none;PADDING-BOTTOM: 0cm;PADDING-TOP: 0cm;PADDING-LEFT: 3.5pt;BORDER-LEFT: none;PADDING-RIGHT: 3.5pt;mso-border-top-alt: solid windowtext .5pt;mso-border-left-alt: none;\" colspan=\"4\">"
                + "                      <p class=\"MsoNormal\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\" align=\"right\">&nbsp;</p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 15.75pt; mso-yfti-irow: 25\">"
                + "                    <td height=\"2%\" colspan=\"3\" valign=\"bottom\" nowrap=\"nowrap\" style=\"BORDER-TOP: medium none; HEIGHT: 15.75pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 317pt; BORDER-BOTTOM: medium none; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: windowtext 1pt solid; PADDING-RIGHT: 3.5pt; mso-border-left-alt: solid windowtext .5pt\">"
                + "                      <p align=\"center\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: center; LINE-HEIGHT: normal\">"
                + "                        <b>"
                + "                          <u>"
                + "                            <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Dados Bancários</span>"
                + "                          </u>"
                + "                        </b>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td valign=\"bottom\" nowrap=\"NOWRAP\" style=\"HEIGHT: 15.75pt; WIDTH: 47pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td style=\"BORDER-TOP: medium none; HEIGHT: 15.75pt; BORDER-RIGHT: black 1pt solid; WIDTH: 220pt; BORDER-BOTTOM: medium none; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: none; PADDING-RIGHT: 3.5pt; mso-border-left-alt: solid windowtext .5pt; mso-border-right-alt: solid black .5pt\" valign=\"bottom\" colspan=\"2\" nowrap=\"nowrap\">"
                + "                      <p align=\"center\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: center; LINE-HEIGHT: normal\">"
                + "                        <b>"
                + "                          <u>"
                + "                            <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Dados para Lançamento</span>"
                + "                          </u>"
                + "                        </b>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 24.75pt; mso-yfti-irow: 26\">"
                + "                    <td width=\"215\" height=\"3%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"BORDER-TOP: medium none; HEIGHT: 24.75pt; BORDER-RIGHT: medium none; WIDTH: 116pt; BORDER-BOTTOM: medium none; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: windowtext 1pt solid; PADDING-RIGHT: 3.5pt; mso-border-left-alt: solid windowtext .5pt\">"
                + "                      <p align=\"right\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Banco:</span>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td width=\"303\" colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"BORDER-TOP: medium none; HEIGHT: 24.75pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 107pt; BORDER-BOTTOM: medium none; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-right-alt: solid windowtext .5pt\">";

            string html6 = "                      <p align=\"left\" class=\"style3 MsoNormal\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\">"
                + "                        <strong>" + banco + "</strong>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td valign=\"bottom\" nowrap=\"NOWRAP\" style=\"HEIGHT: 15.75pt; WIDTH: 47pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td width=\"215\" valign=\"bottom\" nowrap=\"nowrap\" style=\"BORDER-TOP: medium none; HEIGHT: 24.75pt; BORDER-RIGHT: medium none; WIDTH: 116pt; BORDER-BOTTOM: medium none; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: none; PADDING-RIGHT: 3.5pt; mso-border-left-alt: none\">"
                + "                      <p align=\"right\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Fornecedor:</span>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td width=\"303\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"BORDER-TOP: medium none; HEIGHT: 24.75pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 107pt; BORDER-BOTTOM: medium none; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-right-alt: solid windowtext .5pt\">"
                + "                      <p align=\"left\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">1099</span>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 24.95pt; mso-yfti-irow: 27\">"
                + "                    <td width=\"215\" height=\"3%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"BORDER-TOP: medium none; HEIGHT: 24.95pt; BORDER-RIGHT: medium none; WIDTH: 116pt; BORDER-BOTTOM: medium none; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: windowtext 1pt solid; PADDING-RIGHT: 3.5pt; mso-border-left-alt: solid windowtext .5pt\">"
                + "                      <p align=\"right\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Agencia:</span>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td width=\"303\" colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"BORDER-TOP: medium none; HEIGHT: 24.75pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 107pt; BORDER-BOTTOM: medium none; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-right-alt: solid windowtext .5pt\">"
                + "                      <p align=\"left\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\">"
                + "                        <strong>"
                + "                          <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">" + agencia + "</span>"
                + "                        </strong>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td valign=\"bottom\" nowrap=\"NOWRAP\" style=\"HEIGHT: 15pt; WIDTH: 47pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td width=\"215\" valign=\"bottom\" nowrap=\"nowrap\" style=\"BORDER-TOP: medium none; HEIGHT: 24.95pt; BORDER-RIGHT: medium none; WIDTH: 116pt; BORDER-BOTTOM: medium none; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: none; PADDING-RIGHT: 3.5pt; mso-border-left-alt: solid windowtext .5pt\">"
                + "                      <p align=\"right\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Conta Contábil:</span>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td width=\"303\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"BORDER-TOP: medium none; HEIGHT: 24.95pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 107pt; BORDER-BOTTOM: medium none; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-right-alt: solid windowtext .5pt\">"
                + "                      <p align=\"left\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">3431019941100</span>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 24.95pt; mso-yfti-irow: 28\">"
                + "                    <td width=\"215\" height=\"3%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"BORDER-TOP: medium none; HEIGHT: 24.95pt; BORDER-RIGHT: medium none; WIDTH: 116pt; BORDER-BOTTOM: medium none; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: windowtext 1pt solid; PADDING-RIGHT: 3.5pt; mso-border-left-alt: solid windowtext .5pt\">"
                + "                      <p align=\"right\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">C.C.:</span>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td width=\"303\" colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"BORDER-TOP: medium none; HEIGHT: 24.75pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 107pt; BORDER-BOTTOM: medium none; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-right-alt: solid windowtext .5pt\">"
                + "                      <p align=\"left\" class=\"style3 MsoNormal\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\">"
                + "                        <strong>" + contaCorrente + "</strong>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td valign=\"bottom\" nowrap=\"NOWRAP\" style=\"HEIGHT: 24.95pt; WIDTH: 47pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td width=\"215\" valign=\"bottom\" nowrap=\"nowrap\" style=\"BORDER-TOP: medium none; HEIGHT: 24.95pt; BORDER-RIGHT: medium none; WIDTH: 116pt; BORDER-BOTTOM: medium none; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: none; PADDING-RIGHT: 3.5pt; mso-border-left-alt: solid windowtext .5pt\">"
                + "                      <p align=\"right\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: right; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Fluxo de Caixa:</span>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td width=\"303\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"BORDER-TOP: medium none; HEIGHT: 24.95pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 107pt; BORDER-BOTTOM: medium none; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-right-alt: solid windowtext .5pt\">"
                + "                      <p align=\"left\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\">";

            string html7 = "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">03.21.118</span>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 15pt; mso-yfti-irow: 29\">"
                + "                    <td width=\"215\" height=\"2%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"BORDER-TOP: medium none; HEIGHT: 15pt; BORDER-RIGHT: medium none; WIDTH: 116pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: windowtext 1pt solid; PADDING-RIGHT: 3.5pt; mso-border-bottom-alt: solid windowtext .5pt; mso-border-left-alt: solid windowtext .5pt\">"
                + "                      <p class=\"MsoNormal\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\" align=\"left\">&nbsp;</p>"
                + "                    </td>"
                + "                    <td width=\"303\" colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"BORDER-TOP: none; HEIGHT: 15pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 154pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-top-alt: solid windowtext .5pt\">"
                + "                      <p class=\"MsoNormal\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\" align=\"left\">&nbsp;</p>"
                + "                    </td>"
                + "                    <td valign=\"bottom\" nowrap=\"NOWRAP\" style=\"BORDER-TOP: medium none; HEIGHT: 15.75pt; BORDER-RIGHT: medium none; WIDTH: 47pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-bottom-alt: solid windowtext .5pt\">"
                + "                      <p class=\"MsoNormal\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\" align=\"left\">&nbsp;</p>"
                + "                    </td>"
                + "                    <td width=\"215\" valign=\"bottom\" nowrap=\"nowrap\" style=\"BORDER-TOP: medium none; HEIGHT: 15pt; BORDER-RIGHT: medium none; WIDTH: 116pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: none; PADDING-RIGHT: 3.5pt; mso-border-bottom-alt: solid windowtext .5pt; mso-border-left-alt: solid windowtext .5pt\">"
                + "                      <p class=\"MsoNormal\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\" align=\"left\">&nbsp;</p>"
                + "                    </td>"
                + "                    <td width=\"303\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"BORDER-TOP: none; HEIGHT: 15pt; BORDER-RIGHT: windowtext 1pt solid; WIDTH: 154pt; BORDER-BOTTOM: windowtext 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-top-alt: solid windowtext .5pt\">"
                + "                      <p class=\"MsoNormal\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: left; LINE-HEIGHT: normal\" align=\"left\">&nbsp;</p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 70pt;mso-yfti-irow: 30;\">"
                + "                    <td width=\"215\" height=\"2%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 15pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td width=\"303\" colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"HEIGHT: 15pt; WIDTH: 154pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td valign=\"bottom\" nowrap=\"NOWRAP\" style=\"HEIGHT: 15.75pt; WIDTH: 47pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 113pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"215\" nowrap=\"nowrap\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 107pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"303\" nowrap=\"NOWRAP\">      </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 15pt; mso-yfti-irow: 33\">"
                + "                    <td width=\"215\" height=\"2%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 15pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td width=\"303\" colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"HEIGHT: 15pt; WIDTH: 154pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 47pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" nowrap=\"NOWRAP\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 220pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" colspan=\"2\" nowrap=\"nowrap\">"
                + "                      <p align=\"center\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: center; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\"> Aprovação</span>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 15pt; mso-yfti-irow: 34\">"
                + "                    <td width=\"215\" height=\"2%\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 15pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td width=\"303\" colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"HEIGHT: 15pt; WIDTH: 154pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 47pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" nowrap=\"NOWRAP\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 113pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"215\" nowrap=\"nowrap\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 107pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"303\" nowrap=\"NOWRAP\">      </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 15pt; mso-yfti-irow: 35\">"
                + "                    <td width=\"215\" height=\"21\" valign=\"bottom\" nowrap=\"nowrap\" style=\"HEIGHT: 15pt; WIDTH: 116pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td width=\"303\" colspan=\"2\" valign=\"bottom\" nowrap=\"NOWRAP\" style=\"HEIGHT: 15pt; WIDTH: 154pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 47pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" nowrap=\"NOWRAP\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 113pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"215\" nowrap=\"nowrap\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 107pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" width=\"303\" nowrap=\"NOWRAP\">      </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 15pt; mso-yfti-irow: 36\">      </tr>"
                + "                  <tr style=\"HEIGHT: 15.75pt; mso-yfti-irow: 37\">"
                + "                    <td height=\"2%\" colspan=\"3\" valign=\"bottom\" nowrap=\"nowrap\" style=\"BORDER-TOP: windowtext 1pt solid;HEIGHT: 15.75pt;BORDER-RIGHT: medium none;WIDTH: 270pt;BORDER-BOTTOM: medium none;PADDING-BOTTOM: 0cm;PADDING-TOP: 0cm;PADDING-LEFT: 3.5pt;BORDER-LEFT: medium none;PADDING-RIGHT: 3.5pt;mso-border-top-alt: solid windowtext .5pt;\">"
                + "                      <p align=\"center\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: center; LINE-HEIGHT: normal\">"
                + "                        <b>"
                + "                          <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Carolina Mello</span>"
                + "                        </b>"
                + "                      </p>";

            string html8 = "                    </td>"
                + "                    <td style=\"HEIGHT: 15.75pt; WIDTH: 47pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" nowrap=\"NOWRAP\" colspan=\"1\">        </td>"
                + "                    <td style=\"BORDER-TOP: windowtext 1pt solid; HEIGHT: 15.75pt; BORDER-RIGHT: medium none; WIDTH: 220pt; BORDER-BOTTOM: medium none; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; BORDER-LEFT: medium none; PADDING-RIGHT: 3.5pt; mso-border-top-alt: solid windowtext .5pt\" valign=\"bottom\" colspan=\"3\" nowrap=\"nowrap\">"
                + "                      <p align=\"center\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: center; LINE-HEIGHT: normal\">"
                + "                        <b>"
                + "                          <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Francisco Matias Silvano</span>"
                + "                        </b>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                  <tr style=\"HEIGHT: 15pt;mso-yfti-irow: 38;mso-yfti-lastrow: yes;\">"
                + "                    <td height=\"8%\" colspan=\"3\" valign=\"bottom\" nowrap=\"nowrap\" style=\"/* HEIGHT: 10pt; */WIDTH: 270pt;PADDING-BOTTOM: 0cm;PADDING-TOP: 0cm;PADDING-LEFT: 3.5pt;PADDING-RIGHT: 3.5pt;\">"
                + "                      <p align=\"center\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: center; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Gerente SAC</span>"
                + "                      </p>"
                + "                    </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 47pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" nowrap=\"NOWRAP\">        </td>"
                + "                    <td style=\"HEIGHT: 15pt; WIDTH: 220pt; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 3.5pt; PADDING-RIGHT: 3.5pt\" valign=\"bottom\" colspan=\"2\" nowrap=\"nowrap\">"
                + "                      <p align=\"center\" class=\"MsoNormal style3\" style=\"MARGIN-BOTTOM: 0pt; TEXT-ALIGN: center; LINE-HEIGHT: normal\">"
                + "                        <span style=\"COLOR: black; mso-fareast-font-family: &quot;Times New Roman&quot;; mso-bidi-font-family: Arial; mso-bidi-font-size: 12.0pt; mso-fareast-language: PT-BR\">Diretor Administrativo</span>"
                + "                      </p>"
                + "                    </td>"
                + "                  </tr>"
                + "                </tbody>"
                + "              </table>"
                + "            </div>"
                + "          </body>"
                + "        </html>";

            gravarArquivoHTML(inc.ID, html1, html2, html3, html4, html5, html6, html7, html8);
        }

        public static void gravarArquivoHTML(int incidentId, string html1, string html2, string html3, string html4, string html5, string html6, string html7, string html8) {
            string nomeArquivo = escolherNomeArquivo(incidentId);
            if(nomeArquivo != null) {
                StreamWriter outputFile = new StreamWriter(nomeArquivo);
                outputFile.WriteLine(html1);
                outputFile.WriteLine(html2);
                outputFile.WriteLine(html3);
                outputFile.WriteLine(html4);
                outputFile.WriteLine(html5);
                outputFile.WriteLine(html6);
                outputFile.WriteLine(html7);
                outputFile.WriteLine(html8);
                outputFile.Close();
            }
        }

        static string escolherNomeArquivo(int idIncident) {
            Stream myStream;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.FileName = "Ressarcimento Ocorrência " + idIncident;
            saveFileDialog1.Filter = "Arquivo HTML|*.html";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            try {
                if(saveFileDialog1.ShowDialog() == DialogResult.OK) {
                    if((myStream = saveFileDialog1.OpenFile()) != null) {
                        myStream.Close();
                        return saveFileDialog1.FileName;
                    }
                }
            }
            catch(ThreadStateException t) {
                MessageBox.Show("Erro ao salvar arquivo.");
                t.Message.ToString();
            }
            return null;
        }

        public string pegarValorCustomFieldPorNome(string nomeCampo) {
            IIncident inc = (IIncident) RecordContext.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Incident);
            List<ICustomAttribute> i = inc.CustomAttributes.ToList();
            foreach(ICustomAttribute icustomAttribute in i) {
                IGenericField g = icustomAttribute.GenericField;
                string a = g.Name;
                if(a.Equals(nomeCampo)) {
                    return g.DataValue.Value.ToString();
                }
            }
            return "";
        }

        public string pegarValorCustomFieldPorIDString(int id, IIncident inc) {
            foreach(ICfVal icfval in inc.CustomField) {
                if(icfval.CfId == id) {
                    if(icfval.ValStr != null) {
                        return icfval.ValStr;
                    }
                }
            }
            return "";
        }

        public string pegarValorCustomFieldPorIDStringContato(int id, IContact contato) {
            foreach(ICfVal icfval in contato.CustomField) {
                if(icfval.CfId == id) {
                    if(icfval.ValStr != null) {
                        return icfval.ValStr;
                    }
                }
            }
            return "";
        }

        public string pegarValorCustomFieldPorIDInt(int id, IIncident inc) {
            foreach(ICfVal icfval in inc.CustomField) {
                if(icfval.CfId == id) {
                    if(icfval.ValInt != null) {
                        return icfval.ValInt.ToString();
                    }
                }
            }
            return "";
        }

        public DateTime pegarValorCustomFieldPorIDDate(int id, IIncident inc) {
            foreach(ICfVal icfval in inc.CustomField) {
                if(icfval.CfId == id) {
                    if(icfval.ValDate != null) {
                        return (DateTime) icfval.ValDate;
                    }
                }
            }
            return default(DateTime);
        }

        public static String pegarIDOrganizacao(String incidentID) {
            string relacionamentoCliente = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.CO.Cliente FROM Incident WHERE Incident.Id = " + incidentID));
            if(relacionamentoCliente != null && !relacionamentoCliente.Equals("") && !relacionamentoCliente.Equals("N/I")) {
                return relacionamentoCliente;
            }
            return null;
        }

    }

    [AddIn("Workspace Ribbon Button AddIn", Version = "1.0.0.0")]
    public class WorkspaceRibbonButtonFactory : IWorkspaceRibbonButtonFactory {

        #region IWorkspaceRibbonButtonFactory Members
        public IWorkspaceRibbonButton CreateControl(bool inDesignMode, IRecordContext RecordContext) {
            return new WorkspaceRibbonAddIn(inDesignMode, RecordContext);
        }

        public System.Drawing.Image Image32 {
            get {
                return Properties.Resources._32x32;
            }
        }
        #endregion

        #region IFactoryBase Members
        public System.Drawing.Image Image16 {
            get {
                return Properties.Resources._16x16;
            }
        }

        public string Text {
            get {
                return "Relatório de Ressarcimento";
            }
        }

        public string Tooltip {
            get {
                return "Esta função gera um relatório referente aos dados de Ressarcimento deste Incidente.";
            }
        }
        #endregion

        #region IAddInBase Members
        public bool Initialize(IGlobalContext GlobalContext) {
            return true;
        }
        #endregion
    }

}