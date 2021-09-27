using System;
using SAPbobsCOM;
using System.Configuration;
using RestSharp;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Net;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Text;
using System.Xml;
using System.Collections;
using System.Windows.Forms;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;

namespace LocalizacionColombia
{
    class Conexion
    {
        #region Atributos

        /// <summary>
        /// Variable para almacenar la última tabla en la que se ingresó un valor
        /// </summary>
        public static SAPbobsCOM.UserTable TablaInsertarValor;

        /// <summary>
        /// Variable para almacenar estado de creacion tablas y campos
        /// </summary>
        public static int lRetCode;
        public static string sErrMsg;

        /// <summary>
        /// Objeto que permite saber cuando una determinada transacción ha terminado su ejecución
        /// </summary>
        static readonly object padlock = new object();

        public SAPbouiCOM.Application SBO_Application;
        public Company oCompany;
        public string FileLog = "SCL_LOC_LOG";
        public string sessionID;
        private string rutaDocs = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
        //private string ip;
        #endregion

        public void ConCompany(SAPbobsCOM.Company oCom, SAPbouiCOM.Application SBO_App)
        {
            oCompany = oCom;
            SBO_Application = SBO_App;
        }


        public void SetApplication()
        {
            SAPbouiCOM.SboGuiApi SboGuiApi = null;
            string sConnectionString = null;
            SboGuiApi = new SAPbouiCOM.SboGuiApi();
            
            sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));            
            // connect to a running SBO Application
            try
            {
                SboGuiApi.Connect(sConnectionString);
                // get an initialized application object
                SBO_Application = SboGuiApi.GetApplication();                
                SBO_Application.SetStatusBarMessage("Se ha iniciado el addon Localizacion Colombia", SAPbouiCOM.BoMessageTime.bmt_Short, false);

                oCompany = new SAPbobsCOM.Company();
                //get DI company (via UI)
                oCompany = (SAPbobsCOM.Company)SBO_Application.Company.GetDICompany();
            }
            catch (Exception ex)
            { //  Connection failed
                System.Windows.Forms.MessageBox.Show("Error al iniciar el addon Localizacion Colombia" + ex.Message, "Error de conexión",
                    System.Windows.Forms.MessageBoxButtons.OKCancel, System.Windows.Forms.MessageBoxIcon.Error,
                    System.Windows.Forms.MessageBoxDefaultButton.Button1, System.Windows.Forms.MessageBoxOptions.DefaultDesktopOnly);
                System.Environment.Exit(0);
            }
        }
        
        #region Manejo campos y tablas de usuario

        /// <summary>
        /// Método encargado de obtner la version del addon
        /// </summary>
        /// <returns>versión del addon</returns>
        public string GetVersionAddonBD()
        {
            try
            {
                SAPbobsCOM.UserTable tbParametros;
                // Se busca la tabla de la lista de tablas de usuario               
                int i = 0;
                do
                {
                    tbParametros = (SAPbobsCOM.UserTable)oCompany.UserTables.Item(i);
                    i++;
                } while ((tbParametros.TableName != "SCL_LOC_VERSION") && i < oCompany.UserTables.Count);

                if (tbParametros != null && tbParametros.TableName.Equals("SCL_LOC_VERSION") && tbParametros.GetByKey("1"))
                {
                    return tbParametros.Name.ToString();
                }
                return string.Empty;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Creacion de los campos y tablas de usuario
        /// </summary>
        /// <param name="versionNueva">Version del addon</param>
        public void CargaCamposUsuarioDBSAP(string versionNueva)
        {
            //Creacion de tabla de usuario
            CrearTabla("SCL_LOC_VERSION", "Version Localización", BoUTBTableType.bott_NoObject);
            //            CrearTabla("SCL_LOC_CONFIG", "Config Localización", BoUTBTableType.bott_NoObject);
            CrearTabla("SCL_ITM4", "Indicadores de retención permi", BoUTBTableType.bott_MasterDataLines);
            CrearTabla("SCL_CRD4", "Indicadores de AutoRetencion", BoUTBTableType.bott_MasterDataLines);
            //            CrearTabla("SCL_IVA_MAYOR", "IVA mayor costo", BoUTBTableType.bott_NoObject);
            SBO_Application.SetStatusBarMessage("Creando campos de usuario iniciales", SAPbouiCOM.BoMessageTime.bmt_Long, false);
            //Creacion de campos de usuario
            AddFieldsUserTables();

            // Valores por defecto para los parámetros y actualización de versión:
            AgregarValorTablaUsuario("SCL_LOC_VERSION", "", "Code", "1", false);
            AgregarValorTablaUsuario("SCL_LOC_VERSION", "1", "Name", versionNueva, false);

        }

        public void añadirComponentes()
        {
            sessionID = ConexionServiceLayer();
            if (string.IsNullOrEmpty(sessionID)) return;
            readJsonUserTables("JsonFiles/UserTables.json");
            readJsonUserFields("JsonFiles/UserFields.json");
            //Agregado 09 / 16 / 2019
            //addReportsCrystal();            
            //readJsonTransactionCodes();
            readJsonMunicipalities();
            //readJsonQueryCategories();
            //readJsonUserQueries();
            //readJsonFormattedSearches();
            readJsonUserTables("MMagneticos/JsonFiles/UserTables.json");
            readJsonUserFields("MMagneticos/JsonFiles/UserFields.json");
            readJsonConcptsMM();
            readJsonFormatsMM();
            //readJsonUserFields("IVA_Mayor/JsonFiles/UserFields.json");
            SBO_Application.SetStatusBarMessage("Instalación Finalizada", SAPbouiCOM.BoMessageTime.bmt_Short, false);
        }

        /// <summary>
        /// Crea los campos de usuario en la base de datos de SAP. Primero valida si el campo existe, para asegurarse 
        /// de crear el campo y retornar el resultado de la operación
        /// </summary>
        private bool AddFieldsUserTables()
        {
            bool res = true;

            SAPbobsCOM.UserFieldsMD oUserFieldsMD;
            string NameTable;

            try
            {
              
                NameTable = "OADM";
                #region campos OADM

                //oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                //oUserFieldsMD.TableName = NameTable;
                //oUserFieldsMD.Name = "SCL_RutaInf";
                //oUserFieldsMD.Description = "IP Servidor";
                //oUserFieldsMD.DefaultValue = @"10.0.1.5\b1_shf";
                //oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
                //oUserFieldsMD.EditSize = 254;
                //oUserFieldsMD.Add();

                //lRetCode = oUserFieldsMD.Add();
                //if (lRetCode != 0)
                //{
                //    if (lRetCode == -1 || lRetCode == -2035 || lRetCode == -5002)
                //    {
                //        //oCompany.GetLastError(out lRetCode, out sErrMsg);
                //    }
                //    else
                //    {
                //        oCompany.GetLastError(out lRetCode, out sErrMsg);
                //        oUserFieldsMD = null;
                //        GC.Collect();
                //        return false;
                //    }
                //}
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                //oUserFieldsMD = null;
                //GC.Collect();

                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUserFieldsMD.TableName = NameTable;
                oUserFieldsMD.Name = "SCL_RutaSL";
                oUserFieldsMD.Description = "URL";
                oUserFieldsMD.DefaultValue = "http://hanab1:50001";
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
                oUserFieldsMD.EditSize = 100;
                oUserFieldsMD.Add();

                lRetCode = oUserFieldsMD.Add();
                if (lRetCode != 0)
                {
                    if (lRetCode == -1 || lRetCode == -2035 || lRetCode == -5002)
                    {
                        //oCompany.GetLastError(out lRetCode, out sErrMsg);
                    }
                    else
                    {
                        oCompany.GetLastError(out lRetCode, out sErrMsg);
                        oUserFieldsMD = null;
                        GC.Collect();
                        return false;
                    }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();

                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUserFieldsMD.TableName = NameTable;
                oUserFieldsMD.Name = "SCL_UsuarioSL";
                oUserFieldsMD.Description = "Usuario";
                oUserFieldsMD.DefaultValue = "manager";
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
                oUserFieldsMD.EditSize = 20;
                oUserFieldsMD.Add();

                lRetCode = oUserFieldsMD.Add();
                if (lRetCode != 0)
                {
                    if (lRetCode == -1 || lRetCode == -2035 || lRetCode == -5002)
                    {
                        //oCompany.GetLastError(out lRetCode, out sErrMsg);
                    }
                    else
                    {
                        oCompany.GetLastError(out lRetCode, out sErrMsg);
                        oUserFieldsMD = null;
                        GC.Collect();
                        return false;
                    }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();

                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUserFieldsMD.TableName = NameTable;
                oUserFieldsMD.Name = "SCL_ClaveSL";
                oUserFieldsMD.Description = "Contraseña";
                oUserFieldsMD.DefaultValue = null;
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
                oUserFieldsMD.EditSize = 50;
                oUserFieldsMD.Add();

                lRetCode = oUserFieldsMD.Add();
                if (lRetCode != 0)
                {
                    if (lRetCode == -1 || lRetCode == -2035 || lRetCode == -5002)
                    {
                        //oCompany.GetLastError(out lRetCode, out sErrMsg);
                    }
                    else
                    {
                        oCompany.GetLastError(out lRetCode, out sErrMsg);
                        oUserFieldsMD = null;
                        GC.Collect();
                        return false;
                    }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();

                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUserFieldsMD.TableName = NameTable;
                oUserFieldsMD.Name = "SCL_CifradoSL";
                oUserFieldsMD.Description = "Cifrado";
                oUserFieldsMD.DefaultValue = null;
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
                oUserFieldsMD.EditSize = 2;

                oUserFieldsMD.Add();

                lRetCode = oUserFieldsMD.Add();
                if (lRetCode != 0)
                {
                    if (lRetCode == -1 || lRetCode == -2035 || lRetCode == -5002)
                    {
                        //oCompany.GetLastError(out lRetCode, out sErrMsg);
                    }
                    else
                    {
                        oCompany.GetLastError(out lRetCode, out sErrMsg);
                        oUserFieldsMD = null;
                        GC.Collect();
                        return false;
                    }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();

                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUserFieldsMD.TableName = NameTable;
                oUserFieldsMD.Name = "SCL_PrcnCom";
                oUserFieldsMD.Description = "Comisión cirujanos (%)";
                oUserFieldsMD.DefaultValue = "1";
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Percentage;
                oUserFieldsMD.Add();

                lRetCode = oUserFieldsMD.Add();
                if (lRetCode != 0)
                {
                    if (lRetCode == -1 || lRetCode == -2035 || lRetCode == -5002)
                    {
                        //oCompany.GetLastError(out lRetCode, out sErrMsg);
                    }
                    else
                    {
                        oCompany.GetLastError(out lRetCode, out sErrMsg);
                        oUserFieldsMD = null;
                        GC.Collect();
                        return false;
                    }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();

                #endregion campos OADM
                NameTable = "INV1";
                #region campos INV1

                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUserFieldsMD.TableName = NameTable;
                oUserFieldsMD.Name = "SCL_Cod_Ret";
                oUserFieldsMD.Description = "Codigo de Retencion";
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
                oUserFieldsMD.EditSize = 4;

                lRetCode = oUserFieldsMD.Add();
                if (lRetCode != 0)
                {
                    if (lRetCode == -1 || lRetCode == -2035 || lRetCode == -5002)
                    { }
                    else
                    {
                        oCompany.GetLastError(out lRetCode, out sErrMsg);
                        oUserFieldsMD = null;
                        GC.Collect();
                        return false;
                    }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();

                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUserFieldsMD.TableName = NameTable;
                oUserFieldsMD.Name = "SCL_Ret_Val";
                oUserFieldsMD.Description = "Valor Retencion";
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                oUserFieldsMD.SubType = BoFldSubTypes.st_Price;
                //oUserFieldsMD.EditSize = 50;

                lRetCode = oUserFieldsMD.Add();
                if (lRetCode != 0)
                {
                    if (lRetCode == -1 || lRetCode == -2035 || lRetCode == -5002)
                    { }
                    else
                    {
                        oCompany.GetLastError(out lRetCode, out sErrMsg);
                        oUserFieldsMD = null;
                        GC.Collect();
                        return false;
                    }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();

                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUserFieldsMD.TableName = NameTable;
                oUserFieldsMD.Name = "SCL_Ret_Prct";
                oUserFieldsMD.Description = "Porcentaje Retencion";
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Float;
                oUserFieldsMD.SubType = BoFldSubTypes.st_Percentage;
                //oUserFieldsMD.EditSize = 50;

                lRetCode = oUserFieldsMD.Add();
                if (lRetCode != 0)
                {
                    if (lRetCode == -1 || lRetCode == -2035 || lRetCode == -5002)
                    { }
                    else
                    {
                        oCompany.GetLastError(out lRetCode, out sErrMsg);
                        oUserFieldsMD = null;
                        GC.Collect();
                        return false;
                    }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();

                #endregion campos INV1

                NameTable = "OITM";
                #region campos OITM

                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUserFieldsMD.TableName = NameTable;
                oUserFieldsMD.Name = "SCL_WTLiable";
                oUserFieldsMD.Description = "Sujeto a retención";
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
                oUserFieldsMD.EditSize = 1;
                oUserFieldsMD.ValidValues.Value = "Y";
                oUserFieldsMD.ValidValues.Description = "Y";
                oUserFieldsMD.ValidValues.Add();
                oUserFieldsMD.ValidValues.Value = "N";
                oUserFieldsMD.ValidValues.Description = "N";
                oUserFieldsMD.ValidValues.Add();
                oUserFieldsMD.DefaultValue = "N";


                lRetCode = oUserFieldsMD.Add();
                if (lRetCode != 0)
                {
                    if (lRetCode == -1 || lRetCode == -2035 || lRetCode == -5002)
                    {
                        //oCompany.GetLastError(out lRetCode, out sErrMsg);
                    }
                    else
                    {
                        oCompany.GetLastError(out lRetCode, out sErrMsg);
                        oUserFieldsMD = null;
                        GC.Collect();
                        return false;
                    }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();


                #endregion campos OITM
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
            }
            return res;
        }

        /// <summary>
        /// Metodo que crear las tablas de usuario 
        /// </summary>
        /// <param name="tabla">Nombre de la tabla</param>
        /// <param name="descripcion">Descripcion de la tabla</param>
        /// <param name="tipo">Tipo de objeto de la tabla</param>
        private bool CrearTabla(string tabla, string descripcion, BoUTBTableType tipo)
        {
            try
            {
                SAPbobsCOM.UserTablesMD oUsrTble = (UserTablesMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                oUsrTble.TableName = tabla;
                oUsrTble.TableDescription = descripcion;
                oUsrTble.TableType = tipo;
                int retVal = oUsrTble.Add();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUsrTble);
                GC.Collect();
                if (retVal != 0)
                {
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Agrega un valor a la tabla de usuario que indica el parámetro nombreTabla.
        /// </summary>
        /// <param name="nombreTabla">Nombre de la tabla</param>
        /// <param name="Code">Code (PK del registro)</param>
        /// <param name="nombreCampo">Nombre del campo (Si campoUsuario:false puede ser Code o Name)</param>
        /// <param name="valor">Valor a inserar en la tabla</param>
        /// <param name="campoUsuario"></param>
        /// <returns></returns>
        public bool AgregarValorTablaUsuario(string nombreTabla, string Code, string nombreCampo, string valor, bool campoUsuario)
        {
            try
            {
                // Se busca la tabla de la lista de tablas de usuario
                if (TablaInsertarValor == null || !TablaInsertarValor.TableName.Equals(nombreTabla))
                {
                    int i = 0;
                    do
                    {
                        TablaInsertarValor = (SAPbobsCOM.UserTable)oCompany.UserTables.Item(i);
                        i++;
                    } while ((TablaInsertarValor.TableName != nombreTabla) && i < oCompany.UserTables.Count);
                }
                if (TablaInsertarValor != null && TablaInsertarValor.TableName.Equals(nombreTabla))
                {

                    ////Primero se consulta si el Code existe
                    if (TablaInsertarValor.GetByKey(Code) || (nombreCampo.Equals("Code") && TablaInsertarValor.GetByKey(valor)))
                    {
                        // Actualiza
                        oCompany.StartTransaction();
                        if (campoUsuario && string.IsNullOrEmpty(TablaInsertarValor.UserFields.Fields.Item(nombreCampo).Value))
                            TablaInsertarValor.UserFields.Fields.Item(nombreCampo).Value = valor;
                        else if (nombreCampo.Equals("Name"))
                            TablaInsertarValor.Name = valor;
                        int resp = 0;
                        TablaInsertarValor.Update();
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        if (resp == 0)
                            return true;
                    }
                    else
                    {
                        // Agrega
                        oCompany.StartTransaction();
                        if (campoUsuario)
                            TablaInsertarValor.UserFields.Fields.Item(nombreCampo).Value = valor;
                        else if (nombreCampo.Equals("Code"))
                        {
                            TablaInsertarValor.Code = valor;
                            TablaInsertarValor.Name = valor;
                        }
                        int resp = TablaInsertarValor.Add();
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        if (resp == 0)
                            return true;
                    }
                }
                return false;
            }
            catch (Exception)
            {
                if (oCompany.InTransaction)
                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                return false;
            }
        }
        #endregion
        public string ConexionServiceLayer()
        {
            try
            {
                if (string.IsNullOrEmpty(DatosGlobServiceLayer.url) || string.IsNullOrEmpty(DatosGlobServiceLayer.userName) || string.IsNullOrEmpty(DatosGlobServiceLayer.password)) 
                {
                    CredencialesSL();
                }
                var cliente = new RestClient(DatosGlobServiceLayer.url);
                string CompanyDB = oCompany.CompanyDB;
                string Password = DatosGlobServiceLayer.password;
                string UserName = DatosGlobServiceLayer.userName;                
                //var cliente = new RestClient(ConfigurationManager.AppSettings["SLAddress"].ToString());
                //string Password = ConfigurationManager.AppSettings["Password"].ToString();
                //string UserName = ConfigurationManager.AppSettings["UserName"].ToString();                
                var data = new Dictionary<string, string>
            {
                {"CompanyDB", (CompanyDB) },
                { "Password", (Password) },
                {"UserName",  (UserName) }
            };
                var body = JsonConvert.SerializeObject(data);
                var request = new RestRequest("Login", Method.POST);
                request.RequestFormat = DataFormat.Json;
                request.AddParameter("application/json", body, ParameterType.RequestBody);
                RestResponse response = (RestResponse)cliente.Execute(request);
                if (response.StatusCode.Equals(HttpStatusCode.OK) || response.StatusCode.Equals(HttpStatusCode.Created))
                {
                    dynamic dyn = JsonConvert.DeserializeObject(response.Content);
                    foreach (var obj in dyn)
                    {
                        if (obj.Name.Equals("SessionId"))
                            sessionID = obj.Value;
                    }
                }
                else if (response.StatusCode.Equals(HttpStatusCode.NotFound) || response.StatusCode.Equals(HttpStatusCode.Unauthorized) || response.StatusCode.Equals(HttpStatusCode.BadRequest))
                {
                    var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                    var jvalue = (JValue)jobject["error"]["message"]["value"];
                    SBO_Application.SetStatusBarMessage(jvalue.Value + oCompany.GetLastErrorDescription() + "Service Layer", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    escribirLog("ServiceLayer: " + jvalue.Value + oCompany.GetLastErrorDescription());
                }
                else if (response.StatusCode == 0)
                {
                    const string message = "El servicio se encuentra detenido";
                    const string caption = "Servicios Service Layer";
                    var result = MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    escribirLog("ServiceLayer: " + message);

                }
                return sessionID;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Conexion Service Layer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return sessionID;
        }
        
        //--------------------------------------------Prueba SL---------------------------

        public void DesconexionServiceLayer()
        {
            //ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);
            var cliente = new RestClient(DatosGlobServiceLayer.url);
            var request = new RestRequest("Logout", Method.POST);
            request.RequestFormat = DataFormat.Json;
            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            //ServicePointManager.ServerCertificateValidationCallback += new System.Net.Security.RemoteCertificateValidationCallback((sender, certificate, chain, policyErrors) => { return true; });
            RestResponse response = (RestResponse)cliente.Execute(request);
            int status = (int)response.StatusCode;
            if (response.StatusCode.Equals(HttpStatusCode.NoContent))
            {

            }
        }

        public void readJsonUserTables(string route)
        {
            try
            {
                //string outputJSON = File.ReadAllText("JsonFiles/UserTables.json", Encoding.Default);
                string outputJSON = File.ReadAllText(route, Encoding.Default);
                JArray parsedArray = JArray.Parse(outputJSON);
                int cantidad = parsedArray.Count;
                
                dynamic dynJson = JsonConvert.DeserializeObject(outputJSON);
                foreach (var item in dynJson)
                {
                    addUserTables(Convert.ToString(item.TableDescription), Convert.ToString(item.TableName), Convert.ToString(item.TableType));
                }
                DesconexionServiceLayer();
                SBO_Application.StatusBar.SetText("Tablas Creadas", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
                //escribirLog("UserTables: " + ex.Message);
            }
        }
        void addUserTables(string description, string table, string type)
        {

            RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
            RestRequest request = new RestRequest("UserTablesMD", Method.POST);
            var data = new Dictionary<string, string>
                {
                    {"TableDescription", description },
                    { "TableName",  table},
                    {"TableType", type }
                };
            var body = JsonConvert.SerializeObject(data);
            request.RequestFormat = DataFormat.Json;
            request.AddCookie("B1SESSION", sessionID);
            request.AddParameter("application/json", body, ParameterType.RequestBody);
            RestResponse response = (RestResponse)cliente.Execute(request);
            if (!response.StatusCode.Equals(HttpStatusCode.Created))
            {
                response = (RestResponse)cliente.Execute(request);
            }

            int status = (int)response.StatusCode;
            //Console.WriteLine(response.StatusDescription);
            if (response.StatusCode.Equals(HttpStatusCode.Created) || response.StatusCode.Equals(""))
            {
                //AdicionarInfoAlTxt("La tabla " + table + " ya existe  ");
                //Application.SBO_Application.SetStatusBarMessage(" La tabla " +table+ " ya existe  " + SAPCon.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
            else
            {
                var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                var jvalue = (JValue)jobject["error"]["message"]["value"];
                SBO_Application.SetStatusBarMessage(jvalue.Value + oCompany.GetLastErrorDescription() + table, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                escribirLog("UserTables: " + jvalue.Value + oCompany.GetLastErrorDescription() + table);

            }
        }


        public void readJsonUserFields(string route)
        {
            try
            {
                ProgressBar bProgreso = new ProgressBar();
                bProgreso.Style = ProgressBarStyle.Blocks;

                //string outputJSON = File.ReadAllText("JsonFiles/UserFields.json", Encoding.Default);
                string outputJSON = File.ReadAllText(route, Encoding.Default);
                string validValues = "";
                JArray parsedArray = JArray.Parse(outputJSON);
                var bodyField = "";
                string name = "";
                string table1 = "";
                int i = 1;
                int cantidad = parsedArray.Count;
                bProgreso.Maximum = cantidad;
                sessionID = ConexionServiceLayer();
                foreach (JObject parsedObject in parsedArray.Children<JObject>())
                {
                    foreach (JProperty parsedProperty in parsedObject.Properties())
                    {
                        string description = "", table = "", subtype = "", type = "", size = "", mandatory = "", defaultValue = "", linkedSystemObject = "";
                        string tag = parsedProperty.Name;
                        string value = Convert.ToString(parsedProperty.Value);
                        var bodyValues = "";
                        if (tag == "UserFieldsMD")
                        {
                            dynamic dynJson = JsonConvert.DeserializeObject(Convert.ToString(parsedProperty.Value));
                            foreach (var item in dynJson)
                            {
                                string tg = item.Name;
                                switch (tg)
                                {
                                    case "Description":
                                        description = item.Value;
                                        break;
                                    case "Name":
                                        name = item.Value;
                                        break;
                                    case "TableName":
                                        table = item.Value;
                                        break;
                                    case "SubType":
                                        subtype = item.Value;
                                        break;
                                    case "Type":
                                        type = item.Value;
                                        break;
                                    case "Size":
                                        size = item.Value;
                                        break;
                                    case "Mandatory":
                                        mandatory = item.Value;
                                        break;
                                    case "LinkedSystemObject":
                                        linkedSystemObject = item.Value;
                                        break;
                                    case "DefaultValue":
                                        defaultValue = item.Value;
                                        break;
                                }
                            }
                            var data = new Dictionary<string, string>
                        {
                            {"Description", description},
                            {"Name", name },
                            {"SubType", subtype },
                            {"TableName", table },
                            {"Type", type },
                            {"Size", size },
                            {"Mandatory", mandatory },
                            {"DefaultValue", defaultValue },
                            {"LinkedSystemObject", linkedSystemObject}
                        };
                            table1 = table;
                            bodyField = JsonConvert.SerializeObject(data);
                        }
                        else if (tag == "ValidValuesMD")
                        {
                            dynamic dynJson = JsonConvert.DeserializeObject(value);
                            foreach (var item in dynJson)
                            {
                                Dictionary<string, string> Values = new Dictionary<string, string>();
                                Values.Add("Value", Convert.ToString(item.Value));
                                Values.Add("Description", Convert.ToString(item.Description));
                                bodyValues = JsonConvert.SerializeObject(Values);
                                validValues += bodyValues + ",";
                            }
                        }
                    }
                    //   SBO_Application.StatusBar.SetText("Creando campo '"+ name+"' en la tabla "+table1, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    //REFRESCAR SESION EN UN PUNTO
                    if (name.Equals("SCL_BaseMinima"))
                    {
                        DesconexionServiceLayer();
                        sessionID = ConexionServiceLayer();
                    }

                    //if (name.Equals("CreatedBy"))
                    //{
                    //    DesconexionServiceLayer();
                    //    sessionID = ConexionServiceLayer();
                    //}
                    addUserFields(bodyField, validValues, name);
                    validValues = null;
                    bodyField = null;
                    i++;
                    if (bProgreso.Value < bProgreso.Maximum)
                    {
                        bProgreso.Increment(1);
                    }
                    else
                    {
                        bProgreso.Value = bProgreso.Minimum;
                    }

                }
                DesconexionServiceLayer();
                //Console.WriteLine(i + " campos creado");
                SBO_Application.StatusBar.SetText("Campos creados", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }

            catch (Exception ex)
            {
                SBO_Application.MessageBox("Campos\n" + ex.Message);
                // escribirLog("UserFields: " + ex.Message);
            }
        }
        void addUserFields(string bodyField, string validVal, string name)
        {
            try
            {
                RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
                RestRequest request = new RestRequest("UserFieldsMD", Method.POST);
                string validValues = "";
                var body = "";
                if (validVal != null)
                {
                    validValues = "*ValidValuesMD*:[" + validVal + "]";
                    body = (bodyField.ToString().Replace('}', ',').Trim() + "" + validValues.ToString().Replace('*', '"').Trim() + "}").Replace('/', ' ').Trim();
                }
                else
                {
                    body = bodyField.ToString();
                }
                if (name.Equals("Levels"))
                {
                    body = body.ToString().Replace("Size", "EditSize");
                }

                request.AddCookie("B1SESSION", sessionID);
                request.AddParameter("application/json", body, ParameterType.RequestBody);
                RestResponse response = (RestResponse)cliente.Execute(request);
                ///Console.WriteLine(body);
                string status = response.StatusCode.ToString();
                //Console.WriteLine(status);
                if (!response.StatusCode.Equals(HttpStatusCode.Created))
                {
                    response = (RestResponse)cliente.Execute(request);
                }

                if (response.StatusCode.Equals(HttpStatusCode.Created) || response.StatusCode.Equals(""))
                {
                    //AdicionarInfoAlTxt("La tabla " + table + " ya existe  ");
                    //SBO_Application.SetStatusBarMessage(" La tabla " +table+ " ya existe  " + SAPCon.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
                else
                {
                    var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                    var jvalue = (JValue)jobject["error"]["message"]["value"];
                    SBO_Application.SetStatusBarMessage(jvalue.Value + oCompany.GetLastErrorDescription() + " " + name, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    escribirLog("UserFields: " + jvalue.Value + oCompany.GetLastErrorDescription() + " " + name);
                }
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("" + ex.Message);
                escribirLog("UserFields: " + ex.Message);
            }
        }


        public void readJsonTransactionCodes()
        {
            try
            {
                string outputJSON = File.ReadAllText("JsonFiles/TransactionCodes.json", Encoding.Default);
                dynamic dynJson = JsonConvert.DeserializeObject(outputJSON);
                sessionID = ConexionServiceLayer();
                foreach (var item in dynJson)
                {
                    addTransactionCodes(Convert.ToString(item.Code), Convert.ToString(item.Description));
                }
                DesconexionServiceLayer();
                SBO_Application.StatusBar.SetText("Codigos de transaccion añadidos", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
                //escribirLog("TransactionCodes: " + ex.Message);
            }
        }

        void addTransactionCodes(string code, string description)
        {
            RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
            RestRequest request = new RestRequest("TransactionCodes", Method.POST);
            var data = new Dictionary<string, string>
                {
                    {"Code", code },
                    {"Description",  description}
                };
            var body = JsonConvert.SerializeObject(data);
            // Console.WriteLine(body);
            request.RequestFormat = DataFormat.Json;
            request.AddCookie("B1SESSION", sessionID);
            request.AddParameter("application/json", body, ParameterType.RequestBody);
            RestResponse response = (RestResponse)cliente.Execute(request);
            int status = (int)response.StatusCode;
            //Console.WriteLine(response.StatusDescription);
            if (response.StatusCode.Equals(HttpStatusCode.Created) || response.StatusCode.Equals(""))
            {
                //AdicionarInfoAlTxt("El codigo de transaccion " + code + " ya existe");
                //Application.SBO_Application.SetStatusBarMessage(" El codigo de transaccion " + code + " ya existe  " + SAPCon.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
            else
            {
                var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                var jvalue = (JValue)jobject["error"]["message"]["value"];
                SBO_Application.SetStatusBarMessage(jvalue.Value + oCompany.GetLastErrorDescription() + code, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                escribirLog("TransactionCodes: " + jvalue.Value + oCompany.GetLastErrorDescription() + code);
            }
        }
        //------------------------------------
        //public void readJsonDepartments()
        //{
        //    try
        //    {
        //        string inputJSON = File.ReadAllText("JsonFiles/DepartamentosTerceros.json", Encoding.Default);
        //        dynamic dynJson = JsonConvert.DeserializeObject(inputJSON);
        //        sessionID = ConexionServiceLayer();
        //        foreach (var item in dynJson)
        //        {
        //            addDepartments(Convert.ToString(item.Codigo), Convert.ToString(item.Nombre));
        //        }
        //        DesconexionServiceLayer();
        //        SBO_Application.StatusBar.SetText("Departamentos añadidos ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        //    }
        //    catch (Exception ex)
        //    {
        //        SBO_Application.MessageBox("Metodo Departamentos\n" + ex.Message);
        //        escribirLog("Departments: " + ex.Message);
        //    }
        //}
        //void addDepartments(string code, string name)
        //{
        //    //RestClient cliente = new RestClient(ConfigurationManager.AppSettings["SLAddress"].ToString());
        //    RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
        //    RestRequest request = new RestRequest("U_SCL_DEPARTAMENTOS", Method.POST);
        //    var data = new Dictionary<string, string>
        //        {
        //            {"Code", code },
        //            {"Name", name}
        //        };
        //    var body = JsonConvert.SerializeObject(data);
        //    // Console.WriteLine(body);
        //    request.RequestFormat = DataFormat.Json;
        //    request.AddCookie("B1SESSION", sessionID);
        //    request.AddParameter("application/json", body, ParameterType.RequestBody);
        //    RestResponse response = (RestResponse)cliente.Execute(request);
        //    int status = (int)response.StatusCode;
        //    //Console.WriteLine(response.StatusDescription);
        //    if (response.StatusCode.Equals(HttpStatusCode.Created) || response.StatusCode.Equals(""))
        //    {
        //        //AdicionarInfoAlTxt("El departamento " + code + " ya existe");
        //        //Application.SBO_Application.SetStatusBarMessage("El departamento " + code + " ya existe", SAPbouiCOM.BoMessageTime.bmt_Short, false);
        //    }
        //    else
        //    {
        //        var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
        //        var jvalue = (JValue)jobject["error"]["message"]["value"];
        //        //SBO_Application.SetStatusBarMessage(jvalue.Value + oCompany.GetLastErrorDescription() + name, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //        escribirLog("Departments: " + jvalue.Value + oCompany.GetLastErrorDescription() + name);
        //    }
        //}
        //---------------------------------------

        public void readJsonMunicipalities()
        {
            try
            {
                SAPbouiCOM.ProgressBar barraProgreso;
                string inputJSON = File.ReadAllText("JsonFiles/MunicipiosTerceros.json", Encoding.Default);
                dynamic dynJson = JsonConvert.DeserializeObject(inputJSON);
                barraProgreso = SBO_Application.StatusBar.CreateProgressBar("Barra de progreso", 1112, false);
                sessionID = ConexionServiceLayer();
                foreach (var item in dynJson)
                {
                    addMunicipalities(Convert.ToString(item.Codigo), Convert.ToString(item.Nombre), Convert.ToString(item.NombreDepto));
                    barraProgreso.Value += 1;
                }
                barraProgreso.Stop();
                GC.Collect();
                DesconexionServiceLayer();
                SBO_Application.StatusBar.SetText("Municipios Añadidos ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("Metodo Municipios\n" + ex.Message);
                escribirLog("Municipalities: " + ex.Message);
            }
        }
        void addMunicipalities(string code, string name, string nameDpt)
        {
            RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
            RestRequest request = new RestRequest("U_SCL_MUNICIPIOS", Method.POST);
            var data = new Dictionary<string, string>
                {
                    {"Code", code },
                    {"Name", name},
                    {"U_SCL_NomDepto", nameDpt },
                };
            var body = JsonConvert.SerializeObject(data);
            //Console.WriteLine(body);
            request.RequestFormat = DataFormat.Json;
            request.AddCookie("B1SESSION", sessionID);
            request.AddParameter("application/json", body, ParameterType.RequestBody);
            RestResponse response = (RestResponse)cliente.Execute(request);
            int status = (int)response.StatusCode;
            // Console.WriteLine(response.StatusDescription);
            if (response.StatusCode.Equals(HttpStatusCode.Created) || response.StatusCode.Equals(""))
            {

            }
            else
            {
                var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                var jvalue = (JValue)jobject["error"]["message"]["value"];
                //SBO_Application.SetStatusBarMessage(jvalue.Value + oCompany.GetLastErrorDescription() + name, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                escribirLog("Municipalities: " + jvalue.Value + oCompany.GetLastErrorDescription() + " " + name);
            }
        }
        //---------------------------------------------- MM -------------------------------------------
        public void readJsonFormatsMM()
        {
            try
            {
                SAPbouiCOM.ProgressBar barraProgreso;
                string inputJSON = File.ReadAllText("MMagneticos/JsonFiles/FormatosMM.json", Encoding.Default);
                //PROBAR
                //JObject jObj = (JObject)JsonConvert.DeserializeObject(inputJSON);
                //int cant = jObj.Count;
                dynamic dynJson = JsonConvert.DeserializeObject(inputJSON);
               // barraProgreso = SBO_Application.StatusBar.CreateProgressBar("Barra de progreso", cant, false);
                sessionID = ConexionServiceLayer();
                foreach (var item in dynJson)
                {
                    addFormatsMM(Convert.ToString(item.Codigo), Convert.ToString(item.Nombre), Convert.ToString(item.Descripcion), Convert.ToDouble(item.ValCuantia), Convert.ToString(item.Cuantia), Convert.ToString(item.Extraccion), Convert.ToString(item.XML));
                 //   barraProgreso.Value += 1;
                }
                DesconexionServiceLayer();
               // barraProgreso.Stop();
                GC.Collect();
                SBO_Application.StatusBar.SetText("Formatos MM Añadidos ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("readJsonFormatsMM: " + ex.Message);
                escribirLog("readJsonFormatsMM: " + ex.Message);
            }
        }
        void addFormatsMM(string code, string name, string desc, double valCuant, string cuant, string ext, string XML)
        {
            RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
            RestRequest request = new RestRequest("U_SCL_MMFORMATOS", Method.POST);
            //request.Timeout = 30 * 1000;
            var data = new Dictionary<string, object>
                {
                    {"Code", code },
                    {"Name", name},
                    {"U_SCL_DesFrmtoMM", desc },
                    {"U_SCL_VrCFrmtoMM", valCuant},
                    {"U_SCL_CuanFrmtoMM", cuant},
                    {"U_SCL_ExtFrmtoMM", ext},
                    {"U_SCL_XMLFrmtoMM", XML}
                };
            var body = JsonConvert.SerializeObject(data);
            request.RequestFormat = DataFormat.Json;
            request.AddCookie("B1SESSION", sessionID);
            request.AddParameter("application/json", body, ParameterType.RequestBody);
            RestResponse response = (RestResponse)cliente.Execute(request);
            int status = (int)response.StatusCode;
            // Console.WriteLine(response.StatusDescription);
            if (response.StatusCode.Equals(HttpStatusCode.Created) || response.StatusCode.Equals(""))
            {

            }
            else
            {
                var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                var jvalue = (JValue)jobject["error"]["message"]["value"];
                //SBO_Application.SetStatusBarMessage(jvalue.Value + oCompany.GetLastErrorDescription() + name, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                escribirLog("FormatsMM: " + jvalue.Value + oCompany.GetLastErrorDescription() + name);
            }
        }

        //---------------------------------------------- CONCEPTOS MM ----------------------------

        public void readJsonConcptsMM()
        {
            try
            {
                //SAPbouiCOM.ProgressBar barraProgreso;
                string inputJSON = File.ReadAllText("MMagneticos/JsonFiles/ConceptosMM.json", Encoding.Default);
                //PROBAR
                //JObject jObj = (JObject)JsonConvert.DeserializeObject(inputJSON);
                //int cant = jObj.Count;
                dynamic dynJson = JsonConvert.DeserializeObject(inputJSON);
                // barraProgreso = SBO_Application.StatusBar.CreateProgressBar("Barra de progreso", cant, false);
                sessionID = ConexionServiceLayer();
                foreach (var item in dynJson)
                {
                    addConcptsMM(Convert.ToString(item.Codigo), Convert.ToString(item.Nombre), Convert.ToString(item.Descripcion), Convert.ToString(item.AplicaFormato));
                    //   barraProgreso.Value += 1;
                }
                DesconexionServiceLayer();
                SBO_Application.StatusBar.SetText("Conceptos MM Añadidos ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                // barraProgreso.Stop();
                GC.Collect();
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("readJsonConcptsMM: " + ex.Message);
                escribirLog("readJsonConcptsMM: " + ex.Message);
            }
        }
        void addConcptsMM(string code, string name, string desc, string aplFrmto)
        {
            RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
            RestRequest request = new RestRequest("U_SCL_MMCONCEPTOS", Method.POST);
            var data = new Dictionary<string, string>
                {
                    {"Code", code },
                    {"Name", name},
                    {"U_SCL_DesCncptoMM", desc },
                    {"U_SCL_ApFrmtoMM", aplFrmto},
                };
            var body = JsonConvert.SerializeObject(data);
            request.RequestFormat = DataFormat.Json;
            request.AddCookie("B1SESSION", sessionID);
            request.AddParameter("application/json", body, ParameterType.RequestBody);
            RestResponse response = (RestResponse)cliente.Execute(request);
            int status = (int)response.StatusCode;
            if (response.StatusCode.Equals(HttpStatusCode.Created) || response.StatusCode.Equals(""))
            {

            }
            else
            {
                var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                var jvalue = (JValue)jobject["error"]["message"]["value"];
                //SBO_Application.SetStatusBarMessage(jvalue.Value + oCompany.GetLastErrorDescription() + name, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                escribirLog("ConceptosMM: " + jvalue.Value + oCompany.GetLastErrorDescription() + name);
            }
        }

        //---------------------------------------------- MM -------------------------------------------


        public void readJsonQueryCategories()
        {
            try
            {
                string outputJSON = File.ReadAllText("JsonFiles/QueryCategories.json", Encoding.Default);
                dynamic dynJson = JsonConvert.DeserializeObject(outputJSON);
                sessionID = ConexionServiceLayer();
                foreach (var item in dynJson)
                {
                    addQueryCategories(Convert.ToString(item.Name));
                }
                DesconexionServiceLayer();
                SBO_Application.StatusBar.SetText("Categorias creadas", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("Añadir caterogias Consultas\n" + ex.Message);
                escribirLog("QueryCategories: " + ex.Message);
            }
        }
        void addQueryCategories(string name)
        {
            RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
            RestRequest request = new RestRequest("QueryCategories", Method.POST);
            var data = new Dictionary<string, string>
                {
                    {"Name", name },
                };
            var body = JsonConvert.SerializeObject(data);
            //Console.WriteLine(body);
            request.RequestFormat = DataFormat.Json;
            request.AddCookie("B1SESSION", sessionID);
            request.AddParameter("application/json", body, ParameterType.RequestBody);
            RestResponse response = (RestResponse)cliente.Execute(request);
            int status = (int)response.StatusCode;
            if (response.StatusCode.Equals(HttpStatusCode.Created) || response.StatusCode.Equals(""))
            {
                SBO_Application.SetStatusBarMessage(" La categoria  " + name + " fue añadida ", SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
            else
            {
                var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                var jvalue = (JValue)jobject["error"]["message"]["value"];
                SBO_Application.SetStatusBarMessage(jvalue.Value + oCompany.GetLastErrorDescription() + name, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                escribirLog("QueryCategories: " + jvalue.Value + oCompany.GetLastErrorDescription() + name);
            }
        }


        public void readJsonUserQueries()
        {
            try
            {
                string inputJSON = File.ReadAllText("JsonFiles/UserQueries.json", Encoding.Default);
                dynamic dynJson = JsonConvert.DeserializeObject(inputJSON);
                sessionID = ConexionServiceLayer();
                foreach (var item in dynJson)
                {
                    addUserQueries(Convert.ToString(item.QueryName), Convert.ToString(item.QueryCategory), Convert.ToString(item.QueryDescription));
                }
                DesconexionServiceLayer();
                SBO_Application.StatusBar.SetText("Consultas agregadas", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("Agregar consultas de ususario\n" + ex.Message);
                escribirLog("UserQueries: " + ex.Message);
            }
        }
        void addUserQueries(string queryName, string queryCategory, string descripcion)
        //static Boolean addUserQueries(string queryName, string queryCategory, string queryDescription, string sessionID)
        {
            string query = String.Format(Properties.Resources.ResourceManager.GetString(queryName));
            int queryCategoryCode = getIdQueryCategories(queryCategory);
            RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
            RestRequest request = new RestRequest("UserQueries", Method.POST);
            var data = new Dictionary<string, object>
                {
                    {"Query", query },
                    {"QueryCategory",  queryCategoryCode},
                    {"QueryDescription", descripcion }
                };
            var body = JsonConvert.SerializeObject(data);
            //Console.WriteLine(body);
            request.RequestFormat = DataFormat.Json;
            request.AddCookie("B1SESSION", sessionID);
            request.AddParameter("application/json", body, ParameterType.RequestBody);
            RestResponse response = (RestResponse)cliente.Execute(request);
            //Console.WriteLine(response.StatusDescription);
            if (response.StatusCode.Equals(HttpStatusCode.Created) || response.StatusCode.Equals(""))
            {
                //AdicionarInfoAlTxt("La consulta" + queryName + " ya existe");
                //Application.SBO_Application.MessageBox("¡El campo ya existe!");
                SBO_Application.SetStatusBarMessage(" Consulta añadida " + descripcion, SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
            else
            {
                var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                var jvalue = (JValue)jobject["error"]["message"]["value"];
                SBO_Application.SetStatusBarMessage(jvalue.Value + oCompany.GetLastErrorDescription() + queryName, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                escribirLog("UserQueries: " + jvalue.Value + oCompany.GetLastErrorDescription() + queryName);
            }
        }
        int getIdQueryCategories(string nameCategory)
        {
            int categoryID = 0;
            try
            {
                RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
                string req = "QueryCategories?$select=Code&$filter=Name eq '" + nameCategory + "'";
                RestRequest request = new RestRequest(req, Method.GET);
                request.RequestFormat = DataFormat.Json;
                request.AddCookie("B1SESSION", sessionID);
                RestResponse response = (RestResponse)cliente.Execute(request);
                var json = response.Content;
                var result = JsonConvert.DeserializeObject<ODataResponse<GetQueryCategoryID>>(json);
                foreach (var data in result.Value)
                {
                    //Console.WriteLine("ID del campo  " + data.Code);
                    categoryID = data.Code;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage(ex.Message + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                //SBO_Application.MessageBox("Categorias Queries\n" + ex.Message);
                escribirLog("QueryCategories: " + ex.Message);

            }

            return categoryID;
        }


        public void readJsonFormattedSearches()
        {
            try
            {
                SAPbobsCOM.Recordset oRctFormattedData;
                oRctFormattedData = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = Properties.Resources.MaxIndField;
                oRctFormattedData.DoQuery(query);
                int index = Convert.ToInt32(oRctFormattedData.Fields.Item("MAX(IndexID)").Value);
                string inputJSON = File.ReadAllText("JsonFiles/FormattedSearches.json", Encoding.Default);
                dynamic dynJson = JsonConvert.DeserializeObject(inputJSON);
                index += 1;
                sessionID = ConexionServiceLayer();
                foreach (var item in dynJson)
                {
                    addFormattedSearches(Convert.ToString(item.QueryName), Convert.ToInt32(item.FormID), Convert.ToString(item.ItemID), Convert.ToString(item.CollumID), index, Convert.ToString(item.Refresh), Convert.ToString(item.FieldID), Convert.ToString(item.ForceRefresh));
                    index += 1;
                }
                DesconexionServiceLayer();
                SBO_Application.StatusBar.SetText("Busquedas formateadas asignadas ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                //SBO_Application.MessageBox("busquedasFormateadas\n" + ex.Message);
                SBO_Application.SetStatusBarMessage("busquedasFormateadas\n" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                escribirLog("FormattedSearches: " + ex.Message);
            }
        }
        void addFormattedSearches(string queryName, int formID, string itemID, string collumID, int contador, string Refresh, string fieldID, string ForceRefresh)
        {
            int queryID = getIdQuery(queryName);
            RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
            RestRequest request = new RestRequest("FormattedSearches", Method.POST);
            var data = new Dictionary<string, object>
                {
                    {"Action", "bofsaQuery"},
                    {"ColumnID", collumID },
                    {"FormID", formID },
                    {"Index", contador },
                    {"ItemID", itemID },
                    {"QueryID", queryID },
                    {"FieldID", fieldID},
                    {"Refresh", Refresh },
                    {"ForceRefresh", ForceRefresh }
                };
            var body = JsonConvert.SerializeObject(data);
            //Console.WriteLine(body);
            request.RequestFormat = DataFormat.Json;
            request.AddCookie("B1SESSION", sessionID);
            request.AddParameter("application/json", body, ParameterType.RequestBody);
            RestResponse response = (RestResponse)cliente.Execute(request);
            //Console.WriteLine(response.StatusDescription);
            if (response.StatusCode.Equals(HttpStatusCode.Created) || response.StatusCode.Equals(""))
            {

                //AdicionarInfoAlTxt("No se asigno la busqueda formateada " + fielID + " ");
                //Application.SBO_Application.SetStatusBarMessage("No se asigno la busqueda formateada " + fielID + " ", SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
            else
            {
                var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                var jvalue = (JValue)jobject["error"]["message"]["value"];
                //SBO_Application.SetStatusBarMessage(jvalue.Value + oCompany.GetLastErrorDescription() + queryName, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                SBO_Application.SetStatusBarMessage(jvalue.Value + " " + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                escribirLog("FormattedSearches: " + jvalue.Value + " " + oCompany.GetLastErrorDescription());
            }
        }
        int getIdQuery(string queryNam)
        {
            int queryID = 0;
            try
            {
                RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
                RestRequest request = new RestRequest("UserQueries?$select=InternalKey&$filter=QueryDescription eq '" + queryNam + "'", Method.GET);
                request.RequestFormat = DataFormat.Json;
                request.AddCookie("B1SESSION", sessionID);
                RestResponse response = (RestResponse)cliente.Execute(request);
                var json = response.Content;
                var result = JsonConvert.DeserializeObject<ODataResponse<GetQueryID>>(json);
                foreach (var data in result.Value)
                {
                    queryID = data.InternalKey;
                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("ID Query \n" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                //SBO_Application.MessageBox("ID Query \n" + ex.Message);
            }

            return queryID;
        }


        public void addReportsCrystal()
        {
            //AddMenuItems();
            //string sXmlFileName = null;
            //sXmlFileName = AppDomain.CurrentDomain.BaseDirectory;
            //sXmlFileName = System.IO.Directory.GetParent(sXmlFileName).ToString() + @"\XmlFiles\ReportsCrystal.xml";
            //XmlDocument xDoc = new XmlDocument();
            //xDoc.Load(sXmlFileName);
            //XmlNodeList rpts = xDoc.GetElementsByTagName("Reports");
            //XmlNodeList lista = ((XmlElement)rpts[0]).GetElementsByTagName("Report");
            //foreach (XmlElement nodo in lista)
            //{
            //    XmlNodeList locationNames = nodo.GetElementsByTagName("locationName");
            //    XmlNodeList names = nodo.GetElementsByTagName("Name");
            //    string name = names.Item(0).InnerText;
            //    string location = locationNames.Item(0).InnerText;
            //    SAPbobsCOM.ReportLayout oReport;
            //    SAPbobsCOM.ReportLayoutsService oLayoutService;
            //    oLayoutService = (SAPbobsCOM.ReportLayoutsService)oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService);
            //    oReport = (SAPbobsCOM.ReportLayout)oLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayout);
            //    oReport.Name = name;
            //    oReport.TypeCode = "RCRI";
            //    oReport.Author = oCompany.UserName;
            //    oReport.Category = SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal;
            //    string newReportCode = null;
            //    try
            //    {
            //        //SAPbobsCOM.ReportLayoutParams oNewReportParams = oLayoutService.AddReportLayoutToMenu(oReport, "SCL_LOC_COL");
            //        SAPbobsCOM.ReportLayoutParams oNewReportParams = oLayoutService.AddReportLayout(oReport);
            //        newReportCode = oNewReportParams.LayoutCode;
            //    }
            //    catch (System.Exception err)
            //    {
            //        string errMessage = err.Message;
            //        SBO_Application.SetStatusBarMessage("¡Informe " + name + " existente! " + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
            //        escribirLog("ReportsCrystal: " + err.Message + " " + name);
            //        return;
            //    }

            //    //Campo U_SCL_RutaInf en la tabla OADM en el cual se especifica la IP del servidor
            //    //SAPbobsCOM.Recordset oRecordset;
            //    string[] paths = { @"\\" + ip + "", "Addon SCL Colombia", "Informes", location };
            //    string fullPath = Path.Combine(paths);
            //    string rptFilePath = fullPath + ".rpt";
            //    SAPbobsCOM.CompanyService oCompanyService = oCompany.GetCompanyService();
            //    SAPbobsCOM.BlobParams oBlobParams = (SAPbobsCOM.BlobParams)oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams);
            //    oBlobParams.Table = "RDOC";
            //    oBlobParams.Field = "Template";
            //    SAPbobsCOM.BlobTableKeySegment oKeySegment = oBlobParams.BlobTableKeySegments.Add();
            //    oKeySegment.Name = "DocCode";
            //    oKeySegment.Value = newReportCode;
            //    SAPbobsCOM.Blob oBlob = (SAPbobsCOM.Blob)oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlob);

            //    FileStream oFile = new FileStream(rptFilePath, System.IO.FileMode.Open);
            //    int fileSize = (int)oFile.Length;
            //    byte[] buf = new byte[fileSize];
            //    oFile.Read(buf, 0, fileSize);
            //    oFile.Close();

            //    // Convert memory buffer to Base64 string 
            //    oBlob.Content = Convert.ToBase64String(buf, 0, fileSize);
            //    try
            //    {
            //        //Upload Blob to database 
            //        oCompanyService.SetBlob(oBlobParams, oBlob);

            //    }
            //    catch (System.Exception ex)
            //    {
            //        string errmsg = ex.Message;
            //    }
            //}
            SBO_Application.StatusBar.SetText("Informes Crystal Reports añadidos", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        }

        //---------------------------- Añadir Menus -------------------------------------------
        public void AddMenuItems()
        {
            try
            {
                SAPbouiCOM.Menus oMenus = null;
                SAPbouiCOM.MenuItem oMenuItem = null;

                // Get the menus collection from the application
                oMenus = SBO_Application.Menus;

                SAPbouiCOM.MenuCreationParams oCreationPackage = null;
                oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));

                string ip = string.Empty;
                if (!oMenus.Exists("SCL_LOC_COL"))
                {
                    // Get the menu collection of the newly added pop-up item 
                    oMenuItem = null;
                    oMenuItem = SBO_Application.Menus.Item("43531");
                    oMenus = oMenuItem.SubMenus;

                    // Create s sub menu
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                    oCreationPackage.UniqueID = "SCL_LOC_COL";
                    oCreationPackage.String = "SCL Localización Col";
                    oCreationPackage.Position = 1;

                    SAPbobsCOM.Recordset oRecordset;
                    oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string query = String.Format(Properties.Resources.IPServidor);
                    oRecordset.DoQuery(query);
                    ip = oRecordset.Fields.Item("U_SCL_RutaInf").Value;
                    string[] paths = { @"\\" + ip + "", "Addon SCL Colombia", "Iconos", "Bandera" };
                    string fullPath = Path.Combine(paths);
                    oCreationPackage.Image = fullPath + ".jpg";
                    oMenus.AddEx(oCreationPackage);
                }

                if (!oMenus.Exists("CIERREFISCAL"))
                {
                    // Get the menu collection of the newly added pop-up item 
                    //oMenuItem = SBO_Application.Menus.Item("1536");
                    oMenuItem = SBO_Application.Menus.Item("SCL_LOC_COL");
                    oMenus = oMenuItem.SubMenus;
                    // Create sub menu
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "CIERREFISCAL";
                    oCreationPackage.String = "Cierre Fiscal";
                    oCreationPackage.Image = null;
                    oMenus.AddEx(oCreationPackage);
                }

                if (!oMenus.Exists("AsisBalTerceros"))
                {
                    oMenuItem = null;
                    oMenuItem = SBO_Application.Menus.Item("SCL_LOC_COL");
                    oMenus = oMenuItem.SubMenus;
                    //oMenuItem = oMenus.Add("AsisBalTerceros", "Asistente - Balance Terceros", SAPbouiCOM.BoMenuType.mt_STRING, 1);
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "AsisBalTerceros";
                    oCreationPackage.String = "Asistente Balance Terceros";
                    oMenus.AddEx(oCreationPackage);
                }
                if (!oMenus.Exists("SCL_CONFIG"))
                {
                    // Get the menu collection of the newly added pop-up item 
                    oMenuItem = null;
                    oMenuItem = SBO_Application.Menus.Item("8192");
                    oMenus = oMenuItem.SubMenus;

                    // Create s sub menu
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                    oCreationPackage.UniqueID = "SCL_CONFIG";
                    oCreationPackage.String = "SCL Parametrización";
                    oCreationPackage.Position = 1;
                    string[] paths = { @"\\" + ip + "", "b1_shf", "Addon SCL Colombia", "Iconos", "Bandera" };
                    string fullPath = Path.Combine(paths);
                    oCreationPackage.Image = fullPath + ".jpg";
                    oMenus.AddEx(oCreationPackage);
                }
                if (!oMenus.Exists("SCL_MODULOS"))
                {
                    // Get the menu collection of the newly added pop-up item 
                    oMenuItem = null;
                    oMenuItem = SBO_Application.Menus.Item("SCL_CONFIG");
                    oMenus = oMenuItem.SubMenus;

                    // Create s sub menu
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "SCL_MODULOS";
                    oCreationPackage.String = "Modulos localización";
                    oCreationPackage.Position = 1;
                    oMenus.AddEx(oCreationPackage);
                }
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
                escribirLog("AddMenuItems: " + ex.Message);
            }
        }

        public void escribirLog(string cadenalog)
        {
            try
            {
                crearCarpeta();
                string ArchivoLog = FileLog + DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString("00") + DateTime.Today.Day.ToString("00") + " ; " + oCompany.CompanyName + " ; " + oCompany.UserName + ".txt";
                string sPath = (rutaDocs + "\\Logs\\" + ArchivoLog);
                System.IO.StreamWriter file = new System.IO.StreamWriter(sPath, true);
                file.WriteLine(DateTime.Now + " : " + cadenalog);
                file.Close();
                //string ip = string.Empty;
                //SAPbobsCOM.Recordset oRecordset;
                //oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //string query = String.Format(Properties.Resources.IPServidor);
                //oRecordset.DoQuery(query);
                //ip = oRecordset.Fields.Item("U_SCL_RutaInf").Value;
                //string[] paths = { @"\\" + ip + "", "Addon SCL Colombia", "Log", "_" };
                //string fullPath = Path.Combine(paths);
                //string ArchivoLog = FileLog + DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString("00") + DateTime.Today.Day.ToString("00") + " ; " + oCompany.CompanyName + ".txt";
                //string sPath = System.IO.Path.GetDirectoryName(fullPath) + "\\" + (ArchivoLog);

                //// sPath = sPath.Substring(1, sPath.Length - 1);
                //System.IO.StreamWriter file = new System.IO.StreamWriter(sPath, true);
                //file.WriteLine(DateTime.Now + " : " + cadenalog);
                //file.Close();
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
            }
        }

        public void crearCarpeta()
        {
            string ruta = rutaDocs + "\\Logs";
            if (!Directory.Exists(ruta))
            {
                //Console.WriteLine("Creando el directorio: {0}", ruta);
                System.IO.Directory.CreateDirectory(ruta);
            }
            ruta = rutaDocs + "\\Logs\\Iconos";
            if (!Directory.Exists(ruta))
            {
                //Console.WriteLine("Creando el directorio: {0}", ruta);
                System.IO.Directory.CreateDirectory(ruta);
            }
        }

        public void CredencialesSL()
        {
            try
            {
                SAPbobsCOM.Recordset oRecordset;

                //SetApplication();
                oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = String.Format(Properties.Resources.CredencialesSL);
                oRecordset.DoQuery(query);
                DatosGlobServiceLayer.url = oRecordset.Fields.Item("U_SCL_RutaSL").Value;
                DatosGlobServiceLayer.userName = oRecordset.Fields.Item("U_SCL_UsuarioSL").Value;
                DatosGlobServiceLayer.password = Decrypt(oRecordset.Fields.Item("U_SCL_ClaveSL").Value);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("Credenciales SL: " + ex.Message);
            }
        }


        public string Encrypt(string texto)
        {
            try
            {
                /*
                byte[] publicKeySCL = {214,46,220,83,160,73,40,39,201,155,19,202,3,11,191,178,56,
                74,90,36,248,103,18,144,170,163,145,87,54,61,34,220,222,
                207,137,149,173,14,92,120,206,222,158,28,40,24,30,16,155,
                108,128,35,230,118,40,121,113,125,216,130,11,24,90,48,194,
                240,105,44,76,34,57,249,228,125,80,38,9,136,29,117,207,139,
                168,181,85,137,126,10,126,242,120,247,121,8,100,12,201,171,
                38,226,193,180,190,117,177,87,143,242,213,11,44,180,113,93,
                106,99,179,68,175,211,164,116,64,148,226,254,172,147};*/
                string key = "SCLPublicKey"; //llave para encriptar datos

                byte[] keyArray;

                byte[] Arreglo_a_Cifrar = UTF8Encoding.UTF8.GetBytes(texto);

                //Se utilizan las clases de encriptación MD5

                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();

                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));

                hashmd5.Clear();

                //Algoritmo TripleDES
                TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();

                tdes.Key = keyArray;
                tdes.Mode = CipherMode.ECB;
                tdes.Padding = PaddingMode.PKCS7;

                ICryptoTransform cTransform = tdes.CreateEncryptor();

                byte[] ArrayResultado = cTransform.TransformFinalBlock(Arreglo_a_Cifrar, 0, Arreglo_a_Cifrar.Length);

                tdes.Clear();

                //se regresa el resultado en forma de una cadena
                texto = Convert.ToBase64String(ArrayResultado, 0, ArrayResultado.Length);

            }
            catch (Exception ex)
            {
                escribirLog("Encrypt: " + ex.Message);
            }
            return texto;
        }

        public string Decrypt(string textoEncriptado)
        {
            try
            {
                /*
                byte[] publicKeySCL = {214,46,220,83,160,73,40,39,201,155,19,202,3,11,191,178,56,
                74,90,36,248,103,18,144,170,163,145,87,54,61,34,220,222,
                207,137,149,173,14,92,120,206,222,158,28,40,24,30,16,155,
                108,128,35,230,118,40,121,113,125,216,130,11,24,90,48,194,
                240,105,44,76,34,57,249,228,125,80,38,9,136,29,117,207,139,
                168,181,85,137,126,10,126,242,120,247,121,8,100,12,201,171,
                38,226,193,180,190,117,177,87,143,242,213,11,44,180,113,93,
                106,99,179,68,175,211,164,116,64,148,226,254,172,147};*/
                string key = "SCLPublicKey"; //llave para encriptar datos
                byte[] keyArray;
                byte[] Array_a_Descifrar = Convert.FromBase64String(textoEncriptado);

                //algoritmo MD5
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();

                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));

                hashmd5.Clear();

                TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();

                tdes.Key = keyArray;
                tdes.Mode = CipherMode.ECB;
                tdes.Padding = PaddingMode.PKCS7;

                ICryptoTransform cTransform = tdes.CreateDecryptor();

                byte[] resultArray = cTransform.TransformFinalBlock(Array_a_Descifrar, 0, Array_a_Descifrar.Length);

                tdes.Clear();
                textoEncriptado = UTF8Encoding.UTF8.GetString(resultArray);

            }
            catch (Exception ex)
            {
                escribirLog("Decrypt: " + ex.Message);
            }
            return textoEncriptado;
        }
    }
}
