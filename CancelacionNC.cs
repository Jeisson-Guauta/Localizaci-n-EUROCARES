using LocalizacionColombia.AsisReclasificacion;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace LocalizacionColombia
{
    class CancelacionNC
    {

        private SAPbouiCOM.Application SBO_Application;
        private SAPbobsCOM.Company oCompany;
        private string sessionID;
        static Conexion oConnection = new Conexion();

        public CancelacionNC(SAPbobsCOM.Company Con, SAPbouiCOM.Application SBO_App)
        {
            this.oCompany = Con;
            this.SBO_Application = SBO_App;
            serviceLayer();
        }

        public void serviceLayer()
        {
            oConnection.ConCompany(oCompany, SBO_Application);
            sessionID = oConnection.ConexionServiceLayer();
            oConnection.ConCompany(oCompany, SBO_Application);
        }
        public void generarAsientoVentas(ArrayList cntaArt, ArrayList nomCuenta, ArrayList valArt, int docEntry, int transIdOriginal, string sn)
        {
            SAPbobsCOM.Recordset oRecordset;
            SAPbobsCOM.Recordset oRdsImpuesto;
            List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
            Dictionary<string, object> row;

            try
            {
                oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRdsImpuesto = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = Properties.Resources.CntaAnulInv.Replace("%", "V");
                oRecordset.DoQuery(query);
                query = Properties.Resources.CntaAnulIVA.Replace("%", "V");
                oRdsImpuesto.DoQuery(query);
                string cntaAnulInv = oRecordset.Fields.Item("AcctCode").Value.ToString();
                string cntaAnulImp = oRdsImpuesto.Fields.Item("AcctCode").Value.ToString();

                cntaAnulInv = oRecordset.Fields.Item("AcctCode").Value.ToString();
                foreach (var index in Enumerable.Range(0, cntaArt.Count))
                {
                    string cuenta = cntaArt[index].ToString();
                    if (cuenta.StartsWith("A"))
                    {
                        cuenta = cuenta.Substring(1);
                        row = new Dictionary<string, object>();
                        row.Add("AccountCode", cuenta);
                        row.Add("Credit", valArt[index]);
                        row.Add("Debit", 0);
                        row.Add("U_SCL_CodeSN", sn);
                        rows.Add(row);
                        row = new Dictionary<string, object>();
                        row.Add("AccountCode", cntaAnulInv);
                        row.Add("Credit", 0);
                        row.Add("Debit", valArt[index]);
                        row.Add("U_SCL_CodeSN", sn);
                        rows.Add(row);
                    }
                    else
                    {
                        row = new Dictionary<string, object>();
                        row.Add("AccountCode", cuenta);
                        row.Add("Credit", valArt[index]);
                        row.Add("Debit", 0);
                        row.Add("U_SCL_CodeSN", sn);
                        rows.Add(row);
                        row = new Dictionary<string, object>();
                        row.Add("AccountCode", cntaAnulImp);
                        row.Add("Credit", 0);
                        row.Add("Debit", valArt[index]);
                        row.Add("U_SCL_CodeSN", sn);
                        rows.Add(row);
                    }
                }
                JObject jsonObj = new JObject();
                jsonObj.Add("Memo", "Anulacion Nota Credito ventas");
                jsonObj.Add("TransactionCode", "ANCV");
                jsonObj.Add("JournalEntryLines", JArray.FromObject(rows));
                var json = JsonConvert.SerializeObject(jsonObj);
                RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
                RestRequest request = new RestRequest("JournalEntries", Method.POST);
                //ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                //ServicePointManager.ServerCertificateValidationCallback += new System.Net.Security.RemoteCertificateValidationCallback((sender, certificate, chain, policyErrors) => { return true; });
                request.AddCookie("B1SESSION", sessionID);
                request.AddParameter("application/json", json, ParameterType.RequestBody);
                request.AddHeader("Remarks", "Anulacion Nota Credito Ventas");
                //Console.WriteLine(json);
                RestResponse response = (RestResponse)cliente.Execute(request);
                string status = response.StatusDescription;
                //Console.WriteLine(status);
                var res = response.Content;
                if (response.StatusCode.Equals(HttpStatusCode.Created))
                {
                    dynamic dynJson = JsonConvert.DeserializeObject(res);
                    int transaccionID = 0;
                    foreach (var item in dynJson)
                    {
                        if (item.Name == "JdtNum") transaccionID = item.Value;
                    }
                    agregarReferenciaAsiento(transaccionID, transIdOriginal, 0);
                    agregarReferenciaAsiento(docEntry, transaccionID, 1);
                }
                else
                {
                    var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                    var jvalue = (JValue)jobject["error"]["message"]["value"];
                    SBO_Application.SetStatusBarMessage(jvalue.Value + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
                oConnection.escribirLog("generarAsientoVentasNC: " + ex.Message);
            }
        }

        public void generarAsientoCompras(ArrayList cntaArt, ArrayList nomCuenta, ArrayList valArt, int docEntry, int transIdOriginal, string sn)
        {
            SAPbobsCOM.Recordset oRecordset;
            SAPbobsCOM.Recordset oRdsImpuesto;
            List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
            Dictionary<string, object> row;
            Dictionary<string, object> row1;
            Dictionary<string, object> row2;
            try
            {
                oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRdsImpuesto = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = Properties.Resources.CntaAnulInv.Replace("%", "C");
                oRecordset.DoQuery(query);
                query = Properties.Resources.CntaAnulIVA.Replace("%", "C");
                oRdsImpuesto.DoQuery(query);
                string cntaAnulInv = oRecordset.Fields.Item("AcctCode").Value.ToString();
                string cntaAnulImp = oRdsImpuesto.Fields.Item("AcctCode").Value.ToString();

                cntaAnulInv = oRecordset.Fields.Item("AcctCode").Value.ToString();
                foreach (var index in Enumerable.Range(0, cntaArt.Count))
                {
                    string cuenta = cntaArt[index].ToString();
                    row = new Dictionary<string, object>();
                    row.Add("AccountCode", cuenta);
                    row.Add("Credit", 0);
                    row.Add("Debit", valArt[index]);
                    row.Add("U_SCL_CodeSN", sn);
                    rows.Add(row);
                    row = new Dictionary<string, object>();
                    row.Add("AccountCode", cntaAnulImp);
                    row.Add("Credit", valArt[index]);
                    row.Add("Debit", 0);
                    row.Add("U_SCL_CodeSN", sn);
                    rows.Add(row);
                }
                JObject jsonObj = new JObject();
                jsonObj.Add("Memo", "Anulacion Nota Credito Compras");
                jsonObj.Add("TransactionCode", "ANCC");
                jsonObj.Add("JournalEntryLines", JArray.FromObject(rows));
                var json = JsonConvert.SerializeObject(jsonObj);
                RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
                RestRequest request = new RestRequest("JournalEntries", Method.POST);
                //ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                //ServicePointManager.ServerCertificateValidationCallback += new System.Net.Security.RemoteCertificateValidationCallback((sender, certificate, chain, policyErrors) => { return true; });
                request.AddCookie("B1SESSION", sessionID);
                request.AddParameter("application/json", json, ParameterType.RequestBody);
                request.AddHeader("Remarks", "Anulacion Nota Credito Compras");
                //Console.WriteLine(json);
                RestResponse response = (RestResponse)cliente.Execute(request);
                string status = response.StatusDescription;
                //Console.WriteLine(status);
                var res = response.Content;
                if (response.StatusCode.Equals(HttpStatusCode.Created))
                {
                    dynamic dynJson = JsonConvert.DeserializeObject(res);
                    int transaccionID = 0;
                    foreach (var item in dynJson)
                    {
                        if (item.Name == "JdtNum") transaccionID = item.Value;
                    }
                    agregarReferenciaAsiento(transaccionID, transIdOriginal, 0);
                    agregarReferenciaAsiento(docEntry, transaccionID, 2);
                }
                else
                {
                    var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                    var jvalue = (JValue)jobject["error"]["message"]["value"];
                    SBO_Application.SetStatusBarMessage(jvalue.Value + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
                oConnection.escribirLog("generarAsientoComprasNC: " + ex.Message);
            }
        }

        public void agregarReferenciaAsiento(int docEntry, int transaccionID, int id)
        {
            try
            {
                RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
                string req = "JournalEntries({0})";
                req = string.Format(req, transaccionID);
                RestRequest request = new RestRequest(req, RestSharp.Method.PATCH);
                JObject jsonObj = new JObject();

                switch (id)
                {
                    case 0:
                        jsonObj.Add("U_SCL_idAsientoNC", docEntry);
                        break;

                    case 1:
                        jsonObj.Add("U_SCL_DocNCV", docEntry);
                        break;

                    case 2:
                        jsonObj.Add("U_SCL_DocNCC", docEntry);
                        break;

                    default:
                        break;
                }

                var json = JsonConvert.SerializeObject(jsonObj);
                request.RequestFormat = DataFormat.Json;
                request.AddCookie("B1SESSION", sessionID);
                //Console.WriteLine(json);
                request.AddParameter("application/json", json, ParameterType.RequestBody);
                RestResponse response = (RestResponse)cliente.Execute(request);//388
                int status = (int)response.StatusCode;
                //Console.WriteLine(response.Content);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("No se referencio la NC en el asiento \n" + ex.Message);
                oConnection.escribirLog("agregarReferenciaAsientoNC: " + ex.Message);
            }
        }

    }
}
