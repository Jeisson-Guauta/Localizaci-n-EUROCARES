using LocalizacionColombia.Controllers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web;
using System.Xml;
using LocalizacionColombia.AsisReclasificacion;
using LocalizacionColombia.AsisBalTerceros;
using System.Collections;
using LocalizacionColombia.Parametrizacion;
using System.Security.Cryptography;

namespace LocalizacionColombia
{
    class Business
    {
        #region Atributos
        public static SAPbobsCOM.Company oCompany;
        public static CompanyService oCmpSrv;
        public static Application SBO_Application;
        public static Form oForm;
        public static Form frmAsistente;
        public static EditText oEdit;
        public static Item oItem;
        public static StaticText oStatic;
        public static Item oNewItem;
        public static Folder oFolder;
        public static CheckBox oChekBox;
        public static ComboBox oComboBox;
        public static Button oButton;
        public static Matrix oMatrix;
        public static Column oColunm;
        public static Grid oGrid;
        public static string FileLog = "SCL_LOC_LOG";
        public static System.Timers.Timer aTimer, bTimer;
        public static bool flagReSend = true;
        public static bool flagVerifiStatus = true;
        public static int lRetCode;
        public static int contSeries = 1;
        public static string sErrMsg;
        public static string CodigoArticulo;
        public static string CodigoSocio;
        static string rutaDocs = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
        //public static string sessionID;
        public Assistant asistente;
        DatGlobAsistente datosG = new DatGlobAsistente();
        static Conexion oConnection = new Conexion();

        //FRC Atributos nuevos
        string strRuta = System.Windows.Forms.Application.StartupPath; //FRC 20200929 Ruta General de la aplicación

        //int formulario;
        #endregion

        /// <summary>
        /// Inicializa manejo de eventos, timers y cargue de datos  globales
        /// </summary>
        /// <param name="oCmpn">conexion con la compañia</param>
        /// <param name="SBO_App">Interface aplicacion SAP</param>
        public Business(SAPbobsCOM.Company oCmpn, Application SBO_App)
        {
            try
            {
                oCompany = oCmpn;
                SBO_Application = SBO_App;
                AddMenuItems();
                cargarDatosGlobalesSAP();
                //startMonitorSAPB1();
                // ACTULIZA ASIENTO CONTABLES (asignarTerceroAsientos(); actualizarTipoContAsientos();) ...
                //Tareas programadas 
                //AnularNCVentas();
                //AsignarTerceroAsientos();
                //ActualizarTipoContAsientos();

                oConnection.ConCompany(oCompany, SBO_Application);
                SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                SBO_Application.FormDataEvent += new _IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
                SBO_Application.ItemEvent += new _IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                SBO_Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);
                // SBO_Application.LayoutKeyEvent += new SAPbouiCOM._IApplicationEvents_LayoutKeyEventEventHandler(SBO_Application_LayoutKeyEvent);
                escribirLog("Inicio Add-on ");
            }
            catch (Exception ex)
            {
                escribirLog("InicioBusiness: " + ex.Message);
                SBO_Application.SetStatusBarMessage("Exception " + ex.Message, BoMessageTime.bmt_Medium, false);
            }
        }

        //Creacion SubMenu Finanzas
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
                //oMenuItem = SBO_Application.Menus.Item("1536");
                //oMenus = oMenuItem.SubMenus;

                //string sPath = null;
                ////Primer Menu
                //sPath = Application.StartupPath;
                ////sPath = sPath.Remove(sPath.Length - 9, 9);
                //oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                //oCreationPackage.UniqueID = "FE_DIAN";
                //oCreationPackage.String = "Facturacion Electronica";
                //oCreationPackage.Enabled = true;
                ////oCreationPackage.Image = sPath + "\\UI.bmp";
                //oCreationPackage.Position = -1;

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

                    string[] paths = { @"" + rutaDocs + "", "Logs", "Iconos", "Bandera" };
                    string fullPath = Path.Combine(paths);
                    oCreationPackage.Image = fullPath + ".jpg";
                    oMenus.AddEx(oCreationPackage);
                    oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams))); ;
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
                    oCreationPackage.Position = 19;
                    //string[] paths = { @"\\" + ip + "", "b1_shf", "Addon SCL Colombia", "Iconos", "Bandera" };
                    string[] paths = { @"" + rutaDocs + "", "Logs", "Iconos", "Bandera" };
                    string fullPath = Path.Combine(paths);
                    oCreationPackage.Image = fullPath + ".jpg";
                    oMenus.AddEx(oCreationPackage);
                    oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
                }
                //if (!oMenus.Exists("SCL_LOCALIZACION"))
                //{
                //    // Get the menu collection of the newly added pop-up item 
                //    oMenuItem = null;
                //    oMenuItem = SBO_Application.Menus.Item("8192");
                //    oMenus = oMenuItem.SubMenus;

                //    // Create s sub menu
                //    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                //    oCreationPackage.UniqueID = "SCL_LOCALIZACION";
                //    oCreationPackage.String = "Parametrización SCL";
                //    oCreationPackage.Position = 1;
                //    //string[] paths = { @"\\" + ip + "", "b1_shf", "Addon SCL Colombia", "Iconos", "Bandera" };
                //    string[] paths = { @"" + rutaDocs + "", "Logs", "Iconos", "Bandera" };
                //    string fullPath = Path.Combine(paths);
                //    oCreationPackage.Image = fullPath + ".jpg";
                //    oMenus.AddEx(oCreationPackage);
                //    oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
                //}
                if (!oMenus.Exists("SCL_PARAMLOC"))
                {

                    // Get the menu collection of the newly added pop-up item 
                    oMenuItem = null;
                    oMenuItem = SBO_Application.Menus.Item("8192");
                    oMenus = oMenuItem.SubMenus;

                    // Create s sub menu
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "SCL_PARAMLOC";
                    oCreationPackage.String = "Parametrizaciones Iniciales Localización";
                    oCreationPackage.Position = 1;
                    string[] paths = { @"" + rutaDocs + "", "Logs", "Iconos", "Bandera" };
                    string fullPath = Path.Combine(paths);
                    oCreationPackage.Image = fullPath + ".jpg";
                    oMenus.AddEx(oCreationPackage);
                    oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
                }
                //if (!oMenus.Exists("SCL_Legaliza"))
                //{
                //    // Get the menu collection of the newly added pop-up item 
                //    oMenuItem = null;
                //    oMenuItem = SBO_Application.Menus.Item("43520");
                //    oMenus = oMenuItem.SubMenus;

                //    // Create s sub menu
                //    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                //    oCreationPackage.UniqueID = "SCL_Legaliza";
                //    oCreationPackage.String = "Modulo Legalizaciones";
                //    oCreationPackage.Position = 20;
                //    oMenus.AddEx(oCreationPackage);
                //}
                //if (!oMenus.Exists("SCL_Leg_Param"))
                //{
                //    // Get the menu collection of the newly added pop-up item 
                //    oMenuItem = null;
                //    oMenuItem = SBO_Application.Menus.Item("SCL_Legaliza");
                //    oMenus = oMenuItem.SubMenus;

                //    // Create s sub menu
                //    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                //    oCreationPackage.UniqueID = "SCL_Leg_Param";
                //    oCreationPackage.String = "Parametros";
                //    oCreationPackage.Position = 1;
                //    oMenus.AddEx(oCreationPackage);
                //}
                //if (!oMenus.Exists("SCL_Leg_Docum"))
                //{
                //    // Get the menu collection of the newly added pop-up item 
                //    oMenuItem = null;
                //    oMenuItem = SBO_Application.Menus.Item("SCL_Legaliza");
                //    oMenus = oMenuItem.SubMenus;

                //    // Create s sub menu
                //    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                //    oCreationPackage.UniqueID = "SCL_Leg_Docum";
                //    oCreationPackage.String = "Legalización";
                //    oCreationPackage.Position = 2;
                //    oMenus.AddEx(oCreationPackage);
                //}
                if (!oMenus.Exists("SCL_Medios"))
                {
                    // Get the menu collection of the newly added pop-up item 
                    oMenuItem = null;
                    //oMenuItem = SBO_Application.Menus.Item("43520");
                    oMenuItem = SBO_Application.Menus.Item("SCL_LOC_COL");
                    oMenus = oMenuItem.SubMenus;

                    // Create s sub menu
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                    oCreationPackage.UniqueID = "SCL_Medios";
                    oCreationPackage.String = "Modulo Medios Magneticos";
                    oCreationPackage.Position = 20;
                    string[] paths = { @"" + rutaDocs + "", "Logs", "Iconos", "Bandera" };
                    string fullPath = Path.Combine(paths);
                    oCreationPackage.Image = fullPath + ".jpg";
                    oMenus.AddEx(oCreationPackage);
                    oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
                }
                if (!oMenus.Exists("SCL_ParamCntsMM"))
                {
                    // Get the menu collection of the newly added pop-up item 
                    oMenuItem = null;
                    oMenuItem = SBO_Application.Menus.Item("SCL_Medios");
                    oMenus = oMenuItem.SubMenus;

                    // Create s sub menu
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "SCL_ParamCntsMM";
                    oCreationPackage.String = "Parametrización Cuentas";
                    oCreationPackage.Position = 1;
                    oMenus.AddEx(oCreationPackage);
                }
                if (!oMenus.Exists("SCL_GenExtMM"))
                {
                    // Get the menu collection of the newly added pop-up item 
                    oMenuItem = null;
                    oMenuItem = SBO_Application.Menus.Item("SCL_Medios");
                    oMenus = oMenuItem.SubMenus;

                    // Create s sub menu
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "SCL_GenExtMM";
                    oCreationPackage.String = "Generación / Extracción";
                    oCreationPackage.Position = 1;
                    oMenus.AddEx(oCreationPackage);
                }
                if(!oMenus.Exists("SCL_BatchAuto"))
                {
                    // Get the menu collection of the newly added pop-up item 
                    oMenuItem = null;
                    //oMenuItem = SBO_Application.Menus.Item("43520");
                    oMenuItem = SBO_Application.Menus.Item("SCL_LOC_COL");
                    oMenus = oMenuItem.SubMenus;

                    // Create s sub menu
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "SCL_BatchAuto";
                    oCreationPackage.String = "Batch Autorretenciones";
                    oCreationPackage.Position = 21;
                    string[] paths = { @"" + rutaDocs + "", "Logs", "Iconos", "Bandera" };
                    string fullPath = Path.Combine(paths);
                    oCreationPackage.Image = fullPath + ".jpg";
                    oMenus.AddEx(oCreationPackage);
                }
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
                Business.escribirLog("AddMenuItems: " + ex.Message);
            }
        }

        /// <summary>
        /// Captura de eventos del menu
        /// </summary>
        /// <param name="pVal">Objeto del evento</param>
        /// <param name="BubbleEvent">Continua el majeo del evento</param>
        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbobsCOM.Recordset oRS = null;
            string strQry = string.Empty;
		    string sql = string.Empty;					  
            #region Asistente - reclasificación de retenciones
            if (pVal.MenuUID == "ReclasificacionOINV" && (datosG.formType == 133 || datosG.formType == 141) || pVal.MenuUID == "ReclasificacionOPCH" && (datosG.formType == 133 || datosG.formType == 141))
            {
                try
                {
                    asistente = new Assistant(datosG.cardCode, datosG.docNum, datosG.formType, oCompany, SBO_Application);
                    if (frmAsistente.VisibleEx) asistente.cargarDatos();
                    frmAsistente.Visible = true;
                    return;
                }
                catch (Exception ex)
                {
                    frmAsistente = null;
                    SAPbouiCOM.FormCreationParams oCreationParams = null;
                    oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
                    oCreationParams.UniqueID = "AsisReclas";
                    oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
                    frmAsistente = SBO_Application.Forms.AddEx(oCreationParams);
                    frmAsistente.Title = "Asistente - reclasificación de retenciones";
                    frmAsistente.Left = 380;
                    frmAsistente.Top = 15;
                    frmAsistente.Height = 300;
                    frmAsistente.Width = 630;
                    frmAsistente.PaneLevel = 0;
                }
            }
            #endregion
            #region Asistente - Balance de terceros
            // Formulario 
            if ((pVal.MenuUID == "AsisBalTerceros") & (pVal.BeforeAction == false))
            {
                try
                {
                    oForm = SBO_Application.Forms.Item("AsisBalTer");
                    oForm.Visible = true;
                }
                catch
                {
                    oForm = null;
                    SAPbouiCOM.FormCreationParams oCreationParams = null;
                    oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
                    oCreationParams.UniqueID = "AsisBalTer";
                    oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                    oForm = SBO_Application.Forms.AddEx(oCreationParams);
                    oForm.Title = "Asistente - Balance de terceros";
                    //oForm.Left = 300;
                    //oForm.Top = 65;
                    oForm.Height = 541;
                    oForm.Width = 891;
                    oForm.Left = (SBO_Application.Desktop.Width - oForm.Width) / 2;
                    oForm.Top = (SBO_Application.Desktop.Height - oForm.Height) / 2;
                    oForm.PaneLevel = 1;
                    oForm.Visible = true;
                }
            }
            #endregion
            #region Configuracion Modulos
            if ((pVal.MenuUID == "SCL_PARAMLOC") & (pVal.BeforeAction == false))
            {
                try
                {
                    oForm = SBO_Application.Forms.Item("ParamLocalizacion");
                    oForm.Visible = true;

                }
                catch
                {
                    oForm = null;
                    //SAPbouiCOM.FormCreationParams oCreationParams = null;
                    //oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));
                    //oCreationParams.UniqueID = "ParamLocalizacion";
                    //oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.
                    //oForm = SBO_Application.Forms.AddEx(oCreationParams);
                    //oForm.Title = "Parametrizaciones Iniciales Localización";
                    //oForm.Left = 500;
                    //oForm.Top = 65;
                    //oForm.Height = 335;
                    //oForm.Width = 500;
                    //oForm.PaneLevel = 0;
                    ////Parametrizar param = new Parametrizar(oCompany, SBO_Application);
                    ////param.agregarComponentesForm();
                    //oForm.Visible = true;
                    FuncionesGenerales.CargarFormularioXML(SBO_Application, @"Parametrizacion\Forms\ParametrizacionesIni.xml", "SCL_ParamIniLoc");                    
                    
                }
            }
            #endregion
            switch (pVal.MenuUID)
            {
                #region Parametros Legalizaciones
                case "SCL_Leg_Param":
                    if (!pVal.BeforeAction)
                    {
                        try
                        {
                            oForm = SBO_Application.Forms.GetForm("SCL_LegParam", 0);
                        }
                        catch
                        {
                            FuncionesGenerales.CargarFormularioXML(SBO_Application, @"Legalizaciones\Forms\LegParametros.xml", "SCL_LegParam");
                            //FuncionesGenerales.CargarFormularioXML(SBO_Application, strRuta + @"\Forms\LegParametros.xml", "SCL_LegParam");
                            Legalizaciones.Eventos oLegEventos = new Legalizaciones.Eventos();
                            oLegEventos.oApp = SBO_Application;
                            oLegEventos.oComp = oCompany;
                            oLegEventos.CargarInfo(SBO_Application.Forms.GetForm("SCL_LegParam", 0));
                        }
                    }
                    break;
                #endregion
                #region Legalizaciones
                case "SCL_Leg_Docum":
                    if (!pVal.BeforeAction)
                    {
                        try
                        {
                            oForm = SBO_Application.Forms.GetForm("SCL_Legalizacion", 0);
                        }
                        catch
                        {
                            oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                            strQry = "SELECT \"U_Crear\" FROM \"@SCL_LEGPARPERM\" WHERE \"U_NomUsua\" = '" + oCompany.UserName + "'";
                            oRS.DoQuery(strQry);
                            if (oRS.RecordCount > 0)
                            {
                                if (oRS.Fields.Item(0).Value.ToString().Equals("Y"))
                                {
                                    FuncionesGenerales.CargarFormularioXML(SBO_Application, @"Legalizaciones\Forms\LegVentana.xml", "SCL_Legalizacion");
                                    Legalizaciones.Eventos oLegEventos = new Legalizaciones.Eventos();
                                    oLegEventos.oApp = SBO_Application;
                                    oLegEventos.oComp = oCompany;
                                    oLegEventos.CargarInformacionInicial(SBO_Application.Forms.GetForm("SCL_Legalizacion", 0));
                                }
                                else
                                {
                                    SBO_Application.MessageBox("El Usuario " + oCompany.UserName + " No tiene permiso sobre el módulo de Legalizaciones (1)");
                                }
                            }
                            else
                            {
                                SBO_Application.MessageBox("El Usuario " + oCompany.UserName + " No tiene permiso sobre el módulo de Legalizaciones (2)");
                            }
                        }
                    }
                    break;
                #endregion
                #region Parametrización Cuentas MM
                case "SCL_ParamCntsMM":
                    if (!pVal.BeforeAction)
                    {
                        try
                        {
                            oForm = SBO_Application.Forms.GetForm("SCL_ParamCntsMM", 0);
                        }
                        catch
                        {
                            FuncionesGenerales.CargarFormularioXML(SBO_Application, @"MMagneticos\Forms\ParametrizacionCuentas.xml", "SCL_ParamCntsMM");
                            MMagneticos.Eventos EvnParam = new MMagneticos.Eventos();
                            EvnParam.oApp = SBO_Application;
                            EvnParam.oComp = oCompany;
                            EvnParam.CargarInformacionInicial(SBO_Application.Forms.GetForm("SCL_ParamCntsMM", 0));
                        }
                    }
                    break;
                #endregion
                #region Generación Cuentas MM
                case "SCL_GenExtMM":
                    if (!pVal.BeforeAction)
                    {
                        try
                        {
                            oForm = SBO_Application.Forms.GetForm("SCL_GeneracionMM", 0);
                        }
                        catch
                        {
                            FuncionesGenerales.CargarFormularioXML(SBO_Application, @"MMagneticos\Forms\GeneracionMM.xml", "SCL_GeneracionMM");
                            MMagneticos.Eventos EvnGen = new MMagneticos.Eventos();
                            EvnGen.oApp = SBO_Application;
                            EvnGen.oComp = oCompany;
                            EvnGen.CargarInformacionIniGeneracion(SBO_Application.Forms.GetForm("SCL_GeneracionMM", 0));
                        }
                    }
                    break;
                #endregion
                #region Batch Autorretenciones
                case "SCL_BatchAuto":
                    if (!pVal.BeforeAction)
                    {
                        SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        //sql = "SELECT \"U_Crear\" FROM \"@SCL_LEGPARPERM\" WHERE \"U_NomUsua\" = '" + oCompany.UserName + "'";
                        sql = "SELECT 1 FROM \"@SCL_LOC_VERSION\" WHERE \"U_SCL_AutB\" = 'Y'";
                        //sql = OK1.Generic.Helpers.configParteConsulta(this.StrPath + @"AddInQueries\Queries_" + this.TipoServerSQL.Trim() + ".xml", "Batch", "/TQueries/Queries/Query");

                        oRec.DoQuery(sql);

                        if (oRec.RecordCount > 0)
                        {
																					  
												 
						 
							 
						 
							
                            FuncionesGenerales.CargarFormularioXML(SBO_Application, @"BatchAutorretenciones\Forms\BatchAuto.xml", "SCL_BatchAuto");
                            BatchAutorretenciones.Eventos EvnBatch = new BatchAutorretenciones.Eventos();
                            EvnBatch.oApp = SBO_Application;
                            EvnBatch.oComp = oCompany;
                            EvnBatch.CargarInformacionIni(SBO_Application.Forms.GetForm("SCL_BatchAuto", 0));
                        }
                        else
                            SBO_Application.SetStatusBarMessage("Para acceder a esta opción debe habilitarla en la Tabla de Parametrización de las Autorretenciones.", BoMessageTime.bmt_Short, true);
                    }
                    break;
                    #endregion
            }
        }

        /// <summary>
        /// Manejo de eventos documentos SAP (et_FORM_DATA_ADD)
        /// </summary>
        /// <param name="BusinessObjectInfo">Objeto del evento</param>
        /// <param name="BubbleEvent">Continua el majeo del evento</param>
        private void SBO_Application_FormDataEvent(ref BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                oForm = SBO_Application.Forms.ActiveForm;
                #region Form 820
                if (oForm.Type == 820 || oForm.Type == -820)
                {

                }
                #endregion
                #region Form 150 y -150
                if (oForm.Type == 150 || oForm.Type == -150)
                {
                    if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_LOAD && BusinessObjectInfo.ActionSuccess)
                    {
                        bool WTLiable = false;
                        Form form = SBO_Application.Forms.Item(BusinessObjectInfo.FormUID);
                        BusinessObject bisObj = form.BusinessObject;
                        SAPbobsCOM.Items oItems = oCompany.GetBusinessObject(BoObjectTypes.oItems);

                        if (!string.IsNullOrEmpty(BusinessObjectInfo.ObjectKey))
                        {
                            oItems.Browser.GetByKeys(BusinessObjectInfo.ObjectKey);

                            if (!string.IsNullOrEmpty(oItems.UserFields.Fields.Item("U_SCL_WTLiable").Value))
                            {
                                if (oItems.UserFields.Fields.Item("U_SCL_WTLiable").Value == "N")
                                {
                                    WTLiable = false;
                                }
                                else
                                {
                                    WTLiable = true;
                                }
                            }
                            else
                            {
                                WTLiable = false;
                            }
                            oItem = oForm.Items.Item("RET01");
                            oItem.ToPane = 99;
                            oItem.FromPane = 99;
                            oItem.Visible = true;

                            if (WTLiable)
                            {
                                oItem = oForm.Items.Item("RET02");
                                oItem.ToPane = 99;
                                oItem.FromPane = 99;
                                oItem.Visible = true;

                                oItem = oForm.Items.Item("RET03");
                                oItem.ToPane = 99;
                                oItem.FromPane = 99;
                                oItem.Visible = true;
                            }
                            else
                            {
                                oItem = oForm.Items.Item("RET02");
                                oItem.ToPane = 99;
                                oItem.FromPane = 99;
                                oItem.Visible = false;

                                oItem = oForm.Items.Item("RET03");
                                oItem.ToPane = 99;
                                oItem.FromPane = 99;
                                oItem.Visible = false;
                            }

                            oItem = oForm.Items.Item("163");

                            oFolder = ((Folder)(oItem.Specific));
                            oFolder.Select();
                            oForm.PaneLevel = 6;
                        }
                    }
                }
                #endregion
                #region Form 133 y -133
                if (oForm.Type == 133 || oForm.Type == -133)
                {
                    if (BusinessObjectInfo.Type == "13")
                    {
                        //Before Event 
                        if ((BusinessObjectInfo.BeforeAction == false))
                        {
                            try
                            {
                                if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_ADD && BusinessObjectInfo.ActionSuccess)
                                {
                                    CompanyService oCompany;
                                    SeriesService oSeriesService;
                                    Series oSeries;
                                    SeriesParams oSeriesParams;
                                    // get company service
                                    oCompany = Business.oCompany.GetCompanyService();
                                    // get series service
                                    oSeriesService = oCompany.GetBusinessService(ServiceTypes.SeriesService);
                                    // get series params
                                    oSeriesParams = oSeriesService.GetDataInterface(SeriesServiceDataInterfaces.ssdiSeriesParams);
                                    // set the number of an existing series

                                    Form form = SBO_Application.Forms.Item(BusinessObjectInfo.FormUID);
                                    BusinessObject bisObj = form.BusinessObject;
                                    string uid = bisObj.Key;


                                    //Test DI method GetByKeys using key recived from UI (IBusinessObjectInfo.UniqueId) 
                                    SAPbobsCOM.Documents oInvoice = Business.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                                    //oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                                    //Obtener inofrmacion del documento creado
                                    oInvoice.Browser.GetByKeys(BusinessObjectInfo.ObjectKey);
                                    int docEntry = 0;
                                    double valorTotal = 0;

                                    valorTotal = oInvoice.DocTotal;
                                    docEntry = oInvoice.DocEntry;

                                    oSeriesParams.Series = oInvoice.Series;
                                    // get the series
                                    oSeries = oSeriesService.GetSeries(oSeriesParams);

                                    envioDocumento(obtenerConsulta(oSeries.Name, BusinessObjectInfo.Type), docEntry, BusinessObjectInfo.Type, valorTotal);
                                }
                            }
                            catch (Exception ex)
                            {
                                escribirLog("DATA_ADD (133)FacturaDeVenta: " + ex.Message);
                                SBO_Application.MessageBox(ex.Message);
                            }
                        }
                        else
                        {

                        }
                    }
                }
                #endregion
                #region Form 65303 y -65303
                if (oForm.Type == 65303 || oForm.Type == -65303)
                {
                    if (BusinessObjectInfo.Type == "13")
                    {
                        //Before Event 
                        if ((BusinessObjectInfo.BeforeAction == false))
                        {
                            try
                            {
                                if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_ADD && BusinessObjectInfo.ActionSuccess)
                                {
                                    CompanyService oCompany;
                                    SeriesService oSeriesService;
                                    Series oSeries;
                                    SeriesParams oSeriesParams;
                                    // get company service
                                    oCompany = Business.oCompany.GetCompanyService();
                                    // get series service
                                    oSeriesService = oCompany.GetBusinessService(ServiceTypes.SeriesService);
                                    // get series params
                                    oSeriesParams = oSeriesService.GetDataInterface(SeriesServiceDataInterfaces.ssdiSeriesParams);
                                    // set the number of an existing series

                                    Form form = SBO_Application.Forms.Item(BusinessObjectInfo.FormUID);
                                    BusinessObject bisObj = form.BusinessObject;
                                    string uid = bisObj.Key;


                                    //Test DI method GetByKeys using key recived from UI (IBusinessObjectInfo.UniqueId) 
                                    SAPbobsCOM.Documents oInvoice = Business.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                                    //oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                                    //Obtener inofrmacion del documento creado
                                    oInvoice.Browser.GetByKeys(BusinessObjectInfo.ObjectKey);
                                    int docEntry = 0;
                                    double valorTotal = 0;

                                    valorTotal = oInvoice.DocTotal;
                                    docEntry = oInvoice.DocEntry;

                                    oSeriesParams.Series = oInvoice.Series;
                                    // get the series
                                    oSeries = oSeriesService.GetSeries(oSeriesParams);

                                    envioDocumento(obtenerConsulta(oSeries.Name, "65303"), docEntry, BusinessObjectInfo.Type, valorTotal);
                                }
                            }
                            catch (Exception ex)
                            {
                                escribirLog("DATA_ADD (65303)NotaDeDebito: " + ex.Message);
                                SBO_Application.MessageBox(ex.Message);
                            }
                        }
                        else
                        {

                        }
                    }
                }
                #endregion
                #region Form 179 y -179
                if (oForm.Type == 179 || oForm.Type == -179)
                {
                    if (BusinessObjectInfo.Type == "14")
                    {
                        //Before Event 
                        if ((BusinessObjectInfo.BeforeAction == false))
                        {
                            try
                            {
                                if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_ADD && BusinessObjectInfo.ActionSuccess)
                                {
                                    CompanyService oCompany;
                                    SeriesService oSeriesService;
                                    Series oSeries;
                                    SeriesParams oSeriesParams;
                                    // get company service
                                    oCompany = Business.oCompany.GetCompanyService();
                                    // get series service
                                    oSeriesService = oCompany.GetBusinessService(ServiceTypes.SeriesService);
                                    // get series params
                                    oSeriesParams = oSeriesService.GetDataInterface(SeriesServiceDataInterfaces.ssdiSeriesParams);
                                    // set the number of an existing series

                                    Form form = SBO_Application.Forms.Item(BusinessObjectInfo.FormUID);
                                    BusinessObject bisObj = form.BusinessObject;
                                    string uid = bisObj.Key;


                                    //Test DI method GetByKeys using key recived from UI (IBusinessObjectInfo.UniqueId) 
                                    SAPbobsCOM.Documents oCreditNote = Business.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
                                    //oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                                    //Obtener inofrmacion del documento creado
                                    oCreditNote.Browser.GetByKeys(BusinessObjectInfo.ObjectKey);
                                    int docEntry = 0;
                                    double valorTotal = 0;

                                    valorTotal = oCreditNote.DocTotal;
                                    docEntry = oCreditNote.DocEntry;

                                    oSeriesParams.Series = oCreditNote.Series;
                                    // get the series
                                    oSeries = oSeriesService.GetSeries(oSeriesParams);

                                    envioDocumento(obtenerConsulta(oSeries.Name, BusinessObjectInfo.Type), docEntry, BusinessObjectInfo.Type, valorTotal);
                                }
                            }
                            catch (Exception ex)
                            {
                                escribirLog("DATA_ADD:(179)NotaDeCredito:" + ex.Message);
                                SBO_Application.MessageBox(ex.Message);
                            }
                        }
                        else
                        {

                        }
                    }
                }
                #endregion
                switch (BusinessObjectInfo.FormTypeEx) {
                    #region Form Legalizaciones Parametros
                    case "SCL_LegParam":
                    case "SCL_Legalizacion":
                        Legalizaciones.Eventos pLegEventos = new Legalizaciones.Eventos();
                        pLegEventos.oApp = SBO_Application;
                        pLegEventos.oComp = oCompany;
                        pLegEventos.FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                        break;
                    #endregion
                }
            }
            catch (Exception ex)
            {
                escribirLog("FormDataEvent: " + ex.Message);
                //SBO_Application.MessageBox(ex.Message);
            }
        }

        /// <summary>
        /// Manejo de eventos Formulario
        /// </summary>
        /// <param name="FormUID">Identificador del formulario</param>
        /// <param name="pVal">Tipo de evento</param>
        /// <param name="BubbleEvent">Continua el majeo del evento</param>
        private void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE)
            {

                datosG.formType = pVal.FormType;
            }

            try
            {
                string frm = pVal.FormUID;
                switch (pVal.FormTypeEx)
                {
                    case "150":
                        EventFormArticulo(pVal, FormUID);
                        break;
                    case "134":
                        EventFormSocios(pVal, FormUID);
                        break;
                    case "60006":
                        EventFomularioRet(pVal, FormUID);
                        break;
                    case "133":
                        EventFormVentas(pVal, FormUID);
                        break;
                    case "179":
                        EventFormVentas(pVal, FormUID);
                        break;
                    case "60090":
                        EventFormVentas(pVal, FormUID);
                        break;
                    case "65303":
                        EventFormVentas(pVal, FormUID);
                        break;
                    case "141":
                        EventFormCompra(pVal, FormUID);
                        break;
                    case "143":
                        EventFormCompra(pVal, FormUID);
                        break;
                    case "181":
                        EventFormCompra(pVal, FormUID);
                        break;
                    case "65306":
                        EventFormCompra(pVal, FormUID);
                        break;
                    case "136":
                        EventFormDetallesSociedad(pVal, FormUID);//Crear metodo Detalles de sociedad
                        break;
                    case "820":
                        return;
                    //Eventos Eurocares
                    case "25":
                        EventFormSeries(pVal, FormUID);
                        break;
                    case "940":
                        EventFormTransferencia(pVal, FormUID);
                        break;
                    //---------------------------
                    case "SCL_ParamIniLoc":
                        EventFormParamLoc(pVal, FormUID);
                        break;
                    //case "SCL_LegParam":
                    //case "SCL_Legalizacion":
                    //case "SCL_LegAnticipos":
                    //case "SCL_LegRetenciones":
                    //    Legalizaciones.Eventos oLegEventos = new Legalizaciones.Eventos();
                    //    oLegEventos.oApp = SBO_Application;
                    //    oLegEventos.oComp = oCompany;
                    //    oLegEventos.ItemEvent(pVal, FormUID, out BubbleEvent);
                    //    break;
                    case "SCL_ParamCntsMM":
                    case "SCL_CuentasMM":
                    case "SCL_GeneracionMM":
                        MMagneticos.Eventos EvnMM = new MMagneticos.Eventos();
                        EvnMM.oApp = SBO_Application;
                        EvnMM.oComp = oCompany;
                        EvnMM.ItemEvent(pVal, FormUID, out BubbleEvent);
                        break;
                    case "SCL_BatchAuto":
                        BatchAutorretenciones.Eventos EvBatch = new BatchAutorretenciones.Eventos();
                        EvBatch.oApp = SBO_Application;
                        EvBatch.oComp = oCompany;
                        EvBatch.ItemEvent(pVal, FormUID, out BubbleEvent);
                        break;

                    default:
                        //return;
                        break;
                }
            }
            catch (Exception ex)
            {
                Business.escribirLog("ItemEvent: " + ex.Message);
                SBO_Application.MessageBox(ex.Message);
            }


        }

        public void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //int formType = datosG.formType;
            //Condicion validando los formularios --- Nuevo IF

            if (datosG.formType == 133 || datosG.formType == 141)
            {
                SAPbouiCOM.EditText oEdit;
                SAPbouiCOM.Form oForm;
                try
                {
                    if (eventInfo.BeforeAction == true)
                    {
                        int DocNum = 0;
                        string CardCode = null;
                        Dictionary<string, double> Ret = new Dictionary<string, double>();
                        oForm = SBO_Application.Forms.Item(eventInfo.FormUID);

                        //datosG.formType = oForm.Type;
                        oEdit = ((SAPbouiCOM.EditText)(oForm.Items.Item("8").Specific));
                        DocNum = Convert.ToInt32(oEdit.Value);
                        oEdit = ((SAPbouiCOM.EditText)(oForm.Items.Item("4").Specific));
                        CardCode = oEdit.Value;
                        datosG.cardCode = CardCode;
                        datosG.docNum = DocNum;
                        //Console.WriteLine("DocNum = " + DocNum + "\nCardCode = " + CardCode);
                        GC.Collect();
                    }
                }
                catch (Exception ex)
                {
                    // Application.SBO_Application.MessageBox("Item Event \n" + ex.Message);
                }
            }

        }

        /// <summary>
        /// Valida los eventos del formulario de ventas
        /// </summary>
        /// <param name="pVal">Tipo de evento</param>
        /// <param name="formUID">Identificador del formulario</param>
        private void EventFormVentas(ItemEvent pVal, string formUID)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_ITEM_PRESSED:
                        EventFieldVentas(pVal, formUID);
                        break;
                    case BoEventTypes.et_LOST_FOCUS:
                        EventFieldVentas(pVal, formUID);
                        break;
                    case BoEventTypes.et_CLICK:
                        EventFieldVentas(pVal, formUID);
                        break;
                    case BoEventTypes.et_FORM_LOAD:
                        EventFieldVentas(pVal, formUID);
                        break;
                    case BoEventTypes.et_FORM_ACTIVATE:
                        EventFieldVentas(pVal, formUID);
                        break;
                    case BoEventTypes.et_FORM_DEACTIVATE:
                        EventFieldVentas(pVal, formUID);
                        break;
                    case BoEventTypes.et_FORM_CLOSE:
                        EventFieldVentas(pVal, formUID);
                        break;

                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                Business.escribirLog("EventFormVentas: " + ex.Message);
                SBO_Application.MessageBox(ex.Message);
            }
        }

        /// <summary>
        /// Valida los eventos del formulario de compras
        /// </summary>
        /// <param name="pVal">Tipo de evento</param>
        /// <param name="formUID">Identificador del formulario</param>
        private void EventFormCompra(ItemEvent pVal, string formUID)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_ITEM_PRESSED:
                        EventFieldCompras(pVal, formUID);
                        break;
                    case BoEventTypes.et_LOST_FOCUS:
                        EventFieldCompras(pVal, formUID);
                        break;
                    case BoEventTypes.et_CLICK:
                        EventFieldCompras(pVal, formUID);
                        break;
                    case BoEventTypes.et_FORM_LOAD:
                        EventFieldCompras(pVal, formUID);
                        break;
                    case BoEventTypes.et_FORM_ACTIVATE:
                        EventFieldCompras(pVal, formUID);
                        break;
                    case BoEventTypes.et_FORM_CLOSE:
                        EventFieldCompras(pVal, formUID);
                        break;
                    case BoEventTypes.et_FORM_DEACTIVATE:
                        EventFieldCompras(pVal, formUID);
                        break;
                    default:

                        break;
                }
            }
            catch (Exception ex)
            {
                Business.escribirLog("EventFormVentas: " + ex.Message);
                SBO_Application.MessageBox(ex.Message);
            }
        }

        private void EventFormParametrizaciones(ItemEvent pVal, string formUID)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_ITEM_PRESSED:

                        break;
                    case BoEventTypes.et_LOST_FOCUS:

                        break;
                    case BoEventTypes.et_CLICK:

                        break;
                    case BoEventTypes.et_FORM_LOAD:
                        break;

                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                Business.escribirLog("EventFormVentas: " + ex.Message);
                SBO_Application.MessageBox(ex.Message);
            }
        }

        /// <summary>
        /// Valida los eventos del formulario datos maestro de articulo
        /// </summary>
        /// <param name="pVal">Tipo de evento</param>
        /// <param name="formUID">Identificador del formulario</param>
        private void EventFormArticulo(ItemEvent pVal, string formUID)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_FORM_LOAD:
                        CrearFolderFinanzas(pVal, formUID);
                        break;
                    case BoEventTypes.et_ITEM_PRESSED:
                        EventFieldFinanza(pVal, formUID);
                        break;
                    case BoEventTypes.et_CLICK:
                        EventFieldFinanza(pVal, formUID);
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                Business.escribirLog("EventFromArticulo: " + ex.Message);
                SBO_Application.MessageBox(ex.Message);
            }
        }

        private void EventFormSeries(ItemEvent pVal, string formUID)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_KEY_DOWN:
                        EventFieldSeries(pVal, formUID);
                        break;
                    case BoEventTypes.et_FORM_LOAD:
                        EventFieldSeries(pVal, formUID);
                        break;
                    case BoEventTypes.et_GOT_FOCUS:
                        EventFieldSeries(pVal, formUID);
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                Business.escribirLog("EventFormSeries: " + ex.Message);
                SBO_Application.MessageBox(ex.Message);
            }
        }

        private void EventFormTransferencia(ItemEvent pVal, string formUID)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_FORM_LOAD:
                        CrearItemsTS(pVal, formUID);
                        break;

                    case BoEventTypes.et_CLICK:
                        EventFieldTranStock(pVal, formUID);
                        break;
                }
            }
            catch (Exception ex)
            {
                Business.escribirLog("EventFormTransferencia: " + ex.Message);
                SBO_Application.MessageBox(ex.Message);
            }
        }
        /// <summary>
        /// Valida los eventos del formulario datos maestro socio de negocios
        /// </summary>
        /// <param name="pVal">Tipo de evento/<param>
        /// <param name="formUID">Identificador del formulario</param>
        private void EventFormSocios(ItemEvent pVal, string formUID)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_FORM_LOAD:
                        CrearItemsSN(pVal, formUID);
                        CamposExog(pVal, formUID);
                        break;
                    case BoEventTypes.et_CLICK:
                        EventFieldSocio(pVal, formUID);
                        break;
                    case BoEventTypes.et_LOST_FOCUS:
                        EventFieldSocio(pVal, formUID);
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                Business.escribirLog("EventFromSocios: " + ex.Message);
                SBO_Application.MessageBox(ex.Message);
            }
        }

        private void EventFormDetallesSociedad(ItemEvent pVal, string formUID)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_FORM_LOAD:
                        CamposDS(pVal, formUID);
                        break;

                    case BoEventTypes.et_CLICK:
                        EventFieldDS(pVal, formUID);
                        break;

                }
            }
            catch (Exception ex)
            {
                Business.escribirLog("EventFromSocios: " + ex.Message);
                SBO_Application.MessageBox(ex.Message);
            }
        }

        private void CamposExog(ItemEvent pVal, string formUID)
        {
            //throw new NotImplementedException();
            if (pVal.Before_Action == true)
            {
                try
                {
                    // get the event sending form
                    oForm = SBO_Application.Forms.Item(formUID);
                    //---------------------------------------------
                    //Creation Folder Tab
                    // add a new folder item to the form
                    oNewItem = oForm.Items.Add("EXOG001", SAPbouiCOM.BoFormItemTypes.it_FOLDER);

                    oItem = oForm.Items.Item("3");

                    oNewItem.Top = oItem.Top;
                    oNewItem.Height = oItem.Height;
                    oNewItem.Width = oItem.Width;
                    oNewItem.Left = oItem.Left + oItem.Width;
                    oNewItem.Visible = true;
                    oFolder = ((Folder)(oNewItem.Specific));
                    oFolder.Caption = "Exogena";
                    oFolder.GroupWith(oItem.UniqueID);
                    oFolder.AutoPaneSelection = true;
                    oFolder.Pane = 99;

                    //oNewItem = oForm.Items.Add("EXO01", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    //oNewItem.ToPane = 99;
                    //oNewItem.FromPane = 99;
                    //oNewItem.Left = 365;
                    //oNewItem.Top = 224;
                    //oNewItem.Height = 14;
                    //oNewItem.Width = 231;
                    //oNewItem.Visible = false;
                    //oEdit = ((oNewItem.Specific));
                    //oEdit.DataBind.SetBound(true, "OCRD", "U_SCL_Identificacion");

                    //oNewItem = oForm.Items.Add("EXG02", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    //oNewItem.ToPane = 99;
                    //oNewItem.FromPane = 99;
                    //oNewItem.Left = 215;
                    //oNewItem.Top = 224;
                    //oNewItem.Height = 15;
                    //oNewItem.Width = 145;
                    //oNewItem.LinkTo = "EXG01";
                    //oNewItem.Visible = false;
                    //oStatic = ((StaticText)(oNewItem.Specific));
                    //oStatic.Caption = "NIT o C.C.";

                    oNewItem = oForm.Items.Add("EXO03", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 365;
                    oNewItem.Top = 224;
                    oNewItem.Height = 14;
                    oNewItem.Width = 231;
                    oNewItem.Visible = false;
                    oEdit = ((oNewItem.Specific));
                    oEdit.DataBind.SetBound(true, "OCRD", "U_SCL_Apellido1");

                    oNewItem = oForm.Items.Add("EXG04", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 215;
                    oNewItem.Top = 224;
                    oNewItem.Height = 15;
                    oNewItem.Width = 145;
                    oNewItem.LinkTo = "EXG03";
                    oNewItem.Visible = false;
                    oStatic = ((StaticText)(oNewItem.Specific));
                    oStatic.Caption = "Primer Apellido";

                    oNewItem = oForm.Items.Add("EXO05", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 365;
                    oNewItem.Top = 244;
                    oNewItem.Height = 14;
                    oNewItem.Width = 231;
                    oNewItem.Visible = false;
                    oEdit = ((oNewItem.Specific));
                    oEdit.DataBind.SetBound(true, "OCRD", "U_SCL_Apellido2");

                    oNewItem = oForm.Items.Add("EXG06", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 215;
                    oNewItem.Top = 244;
                    oNewItem.Height = 15;
                    oNewItem.Width = 145;
                    oNewItem.LinkTo = "EXG05";
                    oNewItem.Visible = false;
                    oStatic = ((StaticText)(oNewItem.Specific));
                    oStatic.Caption = "Segundo Apellido";

                    oNewItem = oForm.Items.Add("EXO07", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 365;
                    oNewItem.Top = 264;
                    oNewItem.Height = 14;
                    oNewItem.Width = 231;
                    oNewItem.Visible = false;
                    oEdit = ((oNewItem.Specific));
                    oEdit.DataBind.SetBound(true, "OCRD", "U_SCL_Nombre1");

                    oNewItem = oForm.Items.Add("EXG08", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 215;
                    oNewItem.Top = 264;
                    oNewItem.Height = 15;
                    oNewItem.Width = 145;
                    oNewItem.LinkTo = "EXG07";
                    oNewItem.Visible = false;
                    oStatic = ((StaticText)(oNewItem.Specific));
                    oStatic.Caption = "Primer Nombre";

                    oNewItem = oForm.Items.Add("EXO09", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 365;
                    oNewItem.Top = 284;
                    oNewItem.Height = 14;
                    oNewItem.Width = 231;
                    oNewItem.Visible = false;
                    oEdit = ((oNewItem.Specific));
                    oEdit.DataBind.SetBound(true, "OCRD", "U_SCL_NombresAdici");

                    oNewItem = oForm.Items.Add("EXG10", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 215;
                    oNewItem.Top = 284;
                    oNewItem.Height = 15;
                    oNewItem.Width = 145;
                    oNewItem.LinkTo = "EXG09";
                    oNewItem.Visible = false;
                    oStatic = ((StaticText)(oNewItem.Specific));
                    oStatic.Caption = "Nombres Adicionales";

                    oNewItem = oForm.Items.Add("EXO11", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 365;
                    oNewItem.Top = 304;
                    oNewItem.Height = 14;
                    oNewItem.Width = 231;
                    oNewItem.Visible = false;
                    oComboBox = (SAPbouiCOM.ComboBox)oNewItem.Specific;
                    oComboBox.DataBind.SetBound(true, "OCRD", "U_SCL_TipoDoc");

                    oNewItem = oForm.Items.Add("EXG12", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 215;
                    oNewItem.Top = 304;
                    oNewItem.Height = 15;
                    oNewItem.Width = 145;
                    oNewItem.LinkTo = "EXG09";
                    oNewItem.Visible = false;
                    oStatic = ((StaticText)(oNewItem.Specific));
                    oStatic.Caption = "Tipo de Documento";

                    oNewItem = oForm.Items.Add("EXO13", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 365;
                    oNewItem.Top = 324;
                    oNewItem.Height = 14;
                    oNewItem.Width = 231;
                    oNewItem.Visible = false;
                    oComboBox = (SAPbouiCOM.ComboBox)oNewItem.Specific;
                    oComboBox.DataBind.SetBound(true, "OCRD", "U_SCL_TipoPersona");

                    oNewItem = oForm.Items.Add("EXG14", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 215;
                    oNewItem.Top = 324;
                    oNewItem.Height = 15;
                    oNewItem.Width = 145;
                    oNewItem.LinkTo = "EXG13";
                    oNewItem.Visible = false;
                    oStatic = ((StaticText)(oNewItem.Specific));
                    oStatic.Caption = "Tipo de Persona";

                    oNewItem = oForm.Items.Add("EXO15", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 365;
                    oNewItem.Top = 344;
                    oNewItem.Height = 14;
                    oNewItem.Width = 231;
                    oNewItem.Visible = false;
                    oComboBox = (SAPbouiCOM.ComboBox)oNewItem.Specific;
                    oComboBox.DataBind.SetBound(true, "OCRD", "U_SCL_RegTributario");

                    oNewItem = oForm.Items.Add("EXG16", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 215;
                    oNewItem.Top = 344;
                    oNewItem.Height = 15;
                    oNewItem.Width = 145;
                    oNewItem.LinkTo = "EXG15";
                    oNewItem.Visible = false;
                    oStatic = ((StaticText)(oNewItem.Specific));
                    oStatic.Caption = "Regimen Tributario";

                    //oNewItem = oForm.Items.Add("EXO17", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    //oNewItem.ToPane = 99;
                    //oNewItem.FromPane = 99;
                    //oNewItem.Left = 365;
                    //oNewItem.Top = 364;
                    //oNewItem.Height = 14;
                    //oNewItem.Width = 231;
                    //oNewItem.Visible = false;
                    //oEdit = ((oNewItem.Specific));
                    //oEdit.DataBind.SetBound(true, "OCRD", "U_SCL_RazonSocial");

                    //oNewItem = oForm.Items.Add("EXG18", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    //oNewItem.ToPane = 99;
                    //oNewItem.FromPane = 99;
                    //oNewItem.Left = 215;
                    //oNewItem.Top = 364;
                    //oNewItem.Height = 15;
                    //oNewItem.Width = 145;
                    //oNewItem.LinkTo = "EXG17";
                    //oNewItem.Visible = false;
                    //oStatic = ((StaticText)(oNewItem.Specific));
                    //oStatic.Caption = "Razón Social";

                    //oNewItem = oForm.Items.Add("EXO19", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    //oNewItem.ToPane = 99;
                    //oNewItem.FromPane = 99;
                    //oNewItem.Left = 365;
                    //oNewItem.Top = 384;
                    //oNewItem.Height = 14;
                    //oNewItem.Width = 231;
                    //oNewItem.Visible = false;
                    //oEdit = ((oNewItem.Specific));
                    //oEdit.DataBind.SetBound(true, "OCRD", "U_SCL_CodigoMun");

                    //oNewItem = oForm.Items.Add("EXG20", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    //oNewItem.ToPane = 99;
                    //oNewItem.FromPane = 99;
                    //oNewItem.Left = 215;
                    //oNewItem.Top = 384;
                    //oNewItem.Height = 15;
                    //oNewItem.Width = 145;
                    //oNewItem.LinkTo = "EXG19";
                    //oNewItem.Visible = false;
                    //oStatic = ((StaticText)(oNewItem.Specific));
                    //oStatic.Caption = "Municipio MM";

                    //oNewItem = oForm.Items.Add("EXO22", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    //oNewItem.ToPane = 99;
                    //oNewItem.FromPane = 99;
                    //oNewItem.Left = 365;
                    //oNewItem.Top = 368;
                    //oNewItem.Height = 14;
                    //oNewItem.Width = 231;
                    //oNewItem.Visible = false;
                    //oEdit = ((oNewItem.Specific));
                    //oEdit.DataBind.SetBound(true, "OCRD", "U_SCL_PaisDom");

                    //oNewItem = oForm.Items.Add("EXG23", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    //oNewItem.ToPane = 99;
                    //oNewItem.FromPane = 99;
                    //oNewItem.Left = 215;
                    //oNewItem.Top = 368;
                    //oNewItem.Height = 15;
                    //oNewItem.Width = 145;
                    //oNewItem.LinkTo = "EXG22";
                    //oNewItem.Visible = false;
                    //oStatic = ((StaticText)(oNewItem.Specific));
                    //oStatic.Caption = "País Domicilió";

                    //oItem = oForm.Items.Item("3");

                    //oFolder = ((Folder)(oItem.Specific));
                    oFolder.Select();
                    oForm.PaneLevel = 0;
                    oForm.Mode = BoFormMode.fm_OK_MODE;
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("crearFormulario: " + ex.Message);
                }
            }
        }

        private void CamposDS(ItemEvent pVal, string formUID)
        {
            //throw new NotImplementedException();
            if (pVal.Before_Action == true)
            {
                try
                {
                    // get the event sending form
                    oForm = SBO_Application.Forms.Item(formUID);
                    //---------------------------------------------
                    //Creation Folder Tab
                    // add a new folder item to the form
                    oNewItem = oForm.Items.Add("CRED_SL", SAPbouiCOM.BoFormItemTypes.it_FOLDER);

                    oItem = oForm.Items.Item("36");

                    oNewItem.Top = oItem.Top;
                    oNewItem.Height = oItem.Height;
                    oNewItem.Width = oItem.Width;
                    oNewItem.Left = oItem.Left + oItem.Width;
                    oNewItem.Visible = true;
                    oFolder = ((Folder)(oNewItem.Specific));
                    oFolder.Caption = "Localización";
                    oFolder.GroupWith(oItem.UniqueID);
                    oFolder.AutoPaneSelection = true;
                    oFolder.Pane = 99;

                    oNewItem = oForm.Items.Add("CRD01", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 245;
                    oNewItem.Top = 140;
                    oNewItem.Height = 14;
                    oNewItem.Width = 231;
                    oNewItem.Visible = false;
                    oEdit = ((oNewItem.Specific));
                    oEdit.DataBind.SetBound(true, "OADM", "U_SCL_RutaSL");

                    oNewItem = oForm.Items.Add("CRD02", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 100;
                    oNewItem.Top = 140;
                    oNewItem.Height = 15;
                    oNewItem.Width = 145;
                    oNewItem.LinkTo = "CRD01";
                    oNewItem.Visible = false;
                    oStatic = ((StaticText)(oNewItem.Specific));
                    oStatic.Caption = "URL";

                    oNewItem = oForm.Items.Add("CRD03", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 245;
                    oNewItem.Top = 160;
                    oNewItem.Height = 14;
                    oNewItem.Width = 231;
                    oNewItem.Visible = false;
                    oNewItem.Enabled = false;
                    oEdit = ((oNewItem.Specific));
                    oEdit.DataBind.SetBound(true, "OADM", "U_SCL_UsuarioSL");

                    oNewItem = oForm.Items.Add("CRD04", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 100;
                    oNewItem.Top = 160;
                    oNewItem.Height = 15;
                    oNewItem.Width = 145;
                    oNewItem.LinkTo = "CRD03";
                    oNewItem.Visible = false;
                    oStatic = ((StaticText)(oNewItem.Specific));
                    oStatic.Caption = "Usuario";

                    oNewItem = oForm.Items.Add("CRD05", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 245;
                    oNewItem.Top = 180;
                    oNewItem.Height = 14;
                    oNewItem.Width = 231;
                    oNewItem.Visible = false;
                    oEdit = ((oNewItem.Specific));
                    oEdit.IsPassword = true;
                    oEdit.DataBind.SetBound(true, "OADM", "U_SCL_ClaveSL");

                    oNewItem = oForm.Items.Add("CRD06", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 100;
                    oNewItem.Top = 180;
                    oNewItem.Height = 15;
                    oNewItem.Width = 145;
                    oNewItem.LinkTo = "CRD05";
                    oNewItem.Visible = false;
                    oStatic = ((StaticText)(oNewItem.Specific));
                    oStatic.Caption = "Contraseña";

                    oNewItem = oForm.Items.Add("CRD07", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 245;
                    oNewItem.Top = 200;
                    oNewItem.Height = 14;
                    oNewItem.Width = 231;
                    oNewItem.Enabled = false;
                    oNewItem.Visible = true;
                    oEdit = ((oNewItem.Specific));
                    oEdit.DataBind.SetBound(true, "OADM", "U_SCL_CifradoSL");

                    oNewItem = oForm.Items.Add("CRD08", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 100;
                    oNewItem.Top = 200;
                    oNewItem.Height = 15;
                    oNewItem.Width = 145;
                    oNewItem.LinkTo = "CRD07";
                    oNewItem.Visible = false;
                    oStatic = ((StaticText)(oNewItem.Specific));
                    oStatic.Caption = "Cifrado";

                    oNewItem = oForm.Items.Add("CRD09", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 245;
                    oNewItem.Top = 240;
                    oNewItem.Height = 14;
                    oNewItem.Width = 231;
                    oNewItem.Enabled = false;
                    oNewItem.Visible = true;
                    oEdit = ((oNewItem.Specific));
                    oEdit.DataBind.SetBound(true, "OADM", "U_SCL_PrcnCom");

                    oNewItem = oForm.Items.Add("CRD010", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 100;
                    oNewItem.Top = 240;
                    oNewItem.Height = 15;
                    oNewItem.Width = 145;
                    oNewItem.LinkTo = "CRD09";
                    oNewItem.Visible = false;
                    oStatic = ((StaticText)(oNewItem.Specific));
                    oStatic.Caption = "Comisión cirujanos (%)";

                    oItem = oForm.Items.Item("35");

                    oFolder = ((Folder)(oItem.Specific));
                    oFolder.Select();
                    oForm.PaneLevel = 2;
                    oForm.Mode = BoFormMode.fm_OK_MODE;
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("crearFormulario: " + ex.Message);
                }
            }
        }
        /// <summary>
        /// Valida los eventos del formulario retenciones por articulo
        /// </summary>
        /// <param name="pVal">Tipo de evento</param>
        /// <param name="formUID">Identificador del formulario</param>
        private void EventFomularioRet(ItemEvent pVal, string formUID)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_FORM_LOAD:
                        EventFieldRet(pVal, formUID);
                        break;
                    case BoEventTypes.et_FORM_CLOSE:
                        EventFieldRet(pVal, formUID);
                        break;
                    //case BoEventTypes.et_FORM_VISIBLE:
                    //    EventFieldRetItem(pVal, formUID);
                    //    break;
                    case BoEventTypes.et_ITEM_PRESSED:
                        EventFieldRet(pVal, formUID);
                        break;
                    case BoEventTypes.et_CLICK:
                        EventFieldRet(pVal, formUID);
                        break;
                    case BoEventTypes.et_VALIDATE:
                        EventFieldRet(pVal, formUID);
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                Business.escribirLog("EventFomularioRet: " + ex.Message);
                SBO_Application.MessageBox(ex.Message);
            }
        }

        /// <summary>
        /// Valida los eventos parametrizaciones inciales de la localizacion
        /// </summary>
        /// <param name="pVal">Tipo de evento</param>
        /// <param name="formUID">Identificador del formulario</param>
        private void EventFormParamLoc(ItemEvent pVal, string formUID)
        {
            try
            {
                switch (pVal.EventType)
                {
                    //case BoEventTypes.et_FORM_LOAD:
                    //    EventFieldParamLoc(pVal, formUID);
                    //    break;
                    case BoEventTypes.et_CLICK:
                        EventFieldParamLoc(pVal, formUID);
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                Business.escribirLog("EventFieldParamLoc: " + ex.Message);
                SBO_Application.MessageBox(ex.Message);
            }
        }
        /// <summary>
        /// Creacion de folder de Fiananza en dato maestro de articulo
        /// </summary>
        /// <param name="pVal">Tipo de evento</param>
        /// <param name="formUID">Identificador del formulario</param>
        private void CrearFolderFinanzas(ItemEvent pVal, string FormUID)
        {
            if (pVal.Before_Action == true)
            {
                try
                {
                    // get the event sending form
                    oForm = SBO_Application.Forms.Item(FormUID);
                    //---------------------------------------------
                    //Creation Folder Tab
                    // add a new folder item to the form
                    oNewItem = oForm.Items.Add("FINANZA001", SAPbouiCOM.BoFormItemTypes.it_FOLDER);

                    oItem = oForm.Items.Item("163");

                    oNewItem.Top = oItem.Top;
                    oNewItem.Height = oItem.Height;
                    oNewItem.Width = oItem.Width;
                    oNewItem.Left = oItem.Left + oItem.Width;
                    oNewItem.Visible = true;
                    oFolder = ((Folder)(oNewItem.Specific));
                    oFolder.Caption = "Finanzas";
                    oFolder.GroupWith(oItem.UniqueID);
                    oFolder.AutoPaneSelection = true;
                    oFolder.Pane = 99;

                    oNewItem = oForm.Items.Add("RET01", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 315;
                    oNewItem.Top = 168;
                    oNewItem.Height = 14;
                    oNewItem.Width = 231;
                    oNewItem.Visible = false;
                    oChekBox = ((CheckBox)(oNewItem.Specific));
                    oChekBox.Caption = "Sujeto a retención";
                    oChekBox.ValOn = "Y";
                    oChekBox.ValOff = "N";
                    oChekBox.DataBind.SetBound(true, "OITM", "U_SCL_WTLiable");

                    oNewItem = oForm.Items.Add("RET02", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 315;
                    oNewItem.Top = 244;
                    oNewItem.Height = 15;
                    oNewItem.Width = 139;
                    oNewItem.LinkTo = "RET01";
                    oNewItem.Visible = false;
                    oStatic = ((StaticText)(oNewItem.Specific));
                    oStatic.Caption = "Código IVA permitido";

                    oNewItem = oForm.Items.Add("RET03", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    oNewItem.ToPane = 99;
                    oNewItem.FromPane = 99;
                    oNewItem.Left = 465;
                    oNewItem.Top = 244;
                    oNewItem.Height = 15;
                    oNewItem.Width = 22;
                    oNewItem.LinkTo = "RET02";
                    oNewItem.Visible = false;
                    oButton = ((Button)(oNewItem.Specific));
                    oButton.Caption = "...";

                    oItem = oForm.Items.Item("163");

                    oFolder = ((Folder)(oItem.Specific));
                    oFolder.Select();
                    oForm.PaneLevel = 6;
                    oForm.Mode = BoFormMode.fm_OK_MODE;
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("crearFormulario: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Creacion de folder de Fiananza en dato maestro de articulo
        /// </summary>
        /// <param name="pVal">Tipo de evento</param>
        /// <param name="formUID">Identificador del formulario</param>
        private void CrearItemsSN(ItemEvent pVal, string FormUID)
        {
            if (pVal.Before_Action == true)
            {
                try
                {
                    // get the event sending form
                    oForm = SBO_Application.Forms.Item(FormUID);
                    //---------------------------------------------
                    //Creation Folder Tab
                    // add a new folder item to the form
                    //---------------------------------------------

                    /*oForm.DataSources.DBDataSources.Add("OCRD");
                    oNewItem = oForm.Items.Add("EXOGENA001", SAPbouiCOM.BoFormItemTypes.it_FOLDER);

                    oItem = oForm.Items.Item("15");

                    oNewItem.Top = oItem.Top;
                    oNewItem.Height = oItem.Height;
                    oNewItem.Width = oItem.Width;
                    oNewItem.Left = oItem.Left + oItem.Width;
                    oNewItem.Visible = true;
                    oFolder = ((Folder)(oNewItem.Specific));
                    oFolder.Caption = "Campos Exogena";
                    oFolder.GroupWith(oItem.UniqueID);
                    oFolder.AutoPaneSelection = true;
                    oFolder.Pane = 0;

                    oNewItem = oForm.Items.Add("EXO02", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oNewItem.ToPane = 0;
                    oNewItem.FromPane = 0;
                    oNewItem.Left = oItem.Left + oItem.Width;
                    oNewItem.Top = oItem.Top;
                    oNewItem.Height = oItem.Height;
                    oNewItem.Width = oItem.Width;
                    oNewItem.Visible = true;
                    oEdit = ((EditText)(oNewItem.Specific));
                    oEdit.DataBind.SetBound(true, "OCRD", "U_SCL_Identificacion");*/

                    oItem = oForm.Items.Item("258");

                    oNewItem = oForm.Items.Add("RETSN01", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oNewItem.ToPane = oItem.ToPane;
                    oNewItem.FromPane = oItem.FromPane;
                    oNewItem.Left = oItem.Left;
                    oNewItem.Top = oItem.Top + 80;
                    oNewItem.Height = oItem.Height;
                    oNewItem.Width = oItem.Width;
                    oNewItem.Visible = true;
                    oStatic = ((StaticText)(oNewItem.Specific));
                    oStatic.Caption = "Autoretenciones";

                    oItem = oForm.Items.Item("259");

                    oNewItem = oForm.Items.Add("RETSN02", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    oNewItem.ToPane = oItem.ToPane;
                    oNewItem.FromPane = oItem.FromPane;
                    oNewItem.Left = oItem.Left;
                    oNewItem.Top = oItem.Top + 80;
                    oNewItem.Height = oItem.Height;
                    oNewItem.Width = oItem.Width;
                    oNewItem.LinkTo = "RETSN01";
                    oNewItem.Visible = true;
                    oButton = ((Button)(oNewItem.Specific));
                    oButton.Caption = "...";

                    oItem = oForm.Items.Item("3");

                    oFolder = ((Folder)(oItem.Specific));
                    oFolder.Select();
                    oForm.PaneLevel = 1;
                    //oForm.Mode = BoFormMode.fm_OK_MODE;
                }
                catch (Exception ex)
                {
                    //SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("crearFormulario: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Creacion de formulario de Retenciones en dato maestro de articulo
        /// </summary>
        /// <param name="pVal">Tipo de evento</param>
        /// <param name="formUID">Identificador del formulario</param>
        private void CrearFormularioRet(ItemEvent pVal, string FormUID, string ItemCode)
        {
            if (pVal.Before_Action == true)
            {
                try
                {
                    try
                    {
                        oForm = SBO_Application.Forms.Item("RETITEM");
                        oForm.Visible = true;
                    }
                    catch
                    {
                        oForm = null;
                        SAPbouiCOM.FormCreationParams oCreationParams = null;
                        oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

                        oCreationParams.UniqueID = "RETITEM";
                        oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;

                        oForm = SBO_Application.Forms.AddEx(oCreationParams);

                        oItem = oForm.Items.Item("ItemCode");
                        oEdit = oItem.Specific;
                        oEdit.Value = ItemCode;

                        oForm.Title = "Código RI impuesto sobre la renta permitido";
                        oForm.DefButton = "1";
                        oForm.AutoManaged = true;
                        oForm.Left = 506;
                        oForm.Top = 134;
                        oForm.Height = 310;
                        oForm.Width = 343;

                        oForm.Visible = true;
                    }
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("crearFormulario: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Creacion de formulario de Retenciones en dato maestro de articulo
        /// </summary>
        /// <param name="pVal">Tipo de evento</param>
        /// <param name="formUID">Identificador del formulario</param>
        private void CrearFormularioAutoRet(ItemEvent pVal, string FormUID, string CardCode)
        {
            if (pVal.Before_Action == true)
            {
                try
                {
                    try
                    {
                        oForm = SBO_Application.Forms.Item("AUTORET");
                        oForm.Visible = true;
                    }
                    catch
                    {
                        oForm = null;
                        SAPbouiCOM.FormCreationParams oCreationParams = null;
                        oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

                        oCreationParams.UniqueID = "AUTORET";
                        oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;

                        oForm = SBO_Application.Forms.AddEx(oCreationParams);

                        oItem = oForm.Items.Item("CardCode");
                        oEdit = oItem.Specific;
                        oEdit.Value = CardCode;

                        oForm.Title = "Código Autoretenciones";
                        oForm.DefButton = "1";
                        oForm.AutoManaged = true;
                        oForm.Left = 506;
                        oForm.Top = 134;
                        oForm.Height = 310;
                        oForm.Width = 343;

                        oForm.Visible = true;
                    }
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("crearFormulario: " + ex.Message);
                }
            }
        }


        /// <summary>
        /// Creacion de items en formulario retenciones por articulo
        /// </summary>
        /// <param name="pVal">Tipo de evento</param>
        /// <param name="formUID">Identificador del formulario</param>
        private void CrearItemsRet(string FormUID)
        {
            try
            {
                oForm.DataSources.DBDataSources.Add("OWHT");
                oForm.DataSources.UserDataSources.Add("SYS_61", BoDataType.dt_LONG_NUMBER, 4);
                oForm.DataSources.UserDataSources.Add("SYS_62", BoDataType.dt_SHORT_TEXT, 1);
                oForm.DataSources.UserDataSources.Add("SYS_65", BoDataType.dt_LONG_NUMBER, 4);
                oForm.DataSources.UserDataSources.Add("SYS_66", BoDataType.dt_SHORT_TEXT, 10);
                oForm.DataSources.UserDataSources.Add("SYS_67", BoDataType.dt_SHORT_TEXT, 4);

                oItem = oForm.Items.Add("ItemCode", BoFormItemTypes.it_EDIT);
                oItem.Left = 7;
                oItem.Top = 246;
                oItem.Height = 19;
                oItem.Width = 65;
                oItem.Visible = false;
                oEdit = (EditText)oItem.Specific;
                oEdit.Value = "";

                oItem = oForm.Items.Add("1", BoFormItemTypes.it_BUTTON);
                oItem.Left = 7;
                oItem.Top = 246;
                oItem.Height = 19;
                oItem.Width = 65;
                oButton = (Button)oItem.Specific;
                oButton.Caption = "OK";

                oItem = oForm.Items.Add("2", BoFormItemTypes.it_BUTTON);
                oItem.Left = 80;
                oItem.Top = 246;
                oItem.Height = 19;
                oItem.Width = 65;
                oButton = (Button)oItem.Specific;
                oButton.Caption = "Cancelar";

                oItem = oForm.Items.Add("3", BoFormItemTypes.it_MATRIX);
                oItem.Left = 5;
                oItem.Top = 5;
                oItem.Height = 240;
                oItem.Width = 320;
                oMatrix = (Matrix)oItem.Specific;
                oMatrix.SelectionMode = BoMatrixSelect.ms_Auto;

                oColunm = oMatrix.Columns.Add("0", BoFormItemTypes.it_EDIT);
                oColunm.TitleObject.Caption = "#";
                oColunm.Width = 22;
                oColunm.Editable = false;
                oColunm.Visible = true;
                oColunm.AffectsFormMode = true;
                oColunm.DataBind.SetBound(true, "", "SYS_61");

                oColunm = oMatrix.Columns.Add("3", BoFormItemTypes.it_EDIT);
                oColunm.TitleObject.Caption = "Código";
                oColunm.Width = 67;
                oColunm.Editable = false;
                oColunm.Visible = true;
                oColunm.AffectsFormMode = true;
                oColunm.DataBind.SetBound(true, "OWHT", "WTCode");

                oColunm = oMatrix.Columns.Add("1", BoFormItemTypes.it_EDIT);
                oColunm.TitleObject.Caption = "Descripción";
                oColunm.Width = 180;
                oColunm.Editable = false;
                oColunm.Visible = true;
                oColunm.AffectsFormMode = true;
                oColunm.DataBind.SetBound(true, "OWHT", "WTName");

                oColunm = oMatrix.Columns.Add("2", BoFormItemTypes.it_CHECK_BOX);
                oColunm.TitleObject.Caption = "Seleccionar";
                oColunm.Width = 33;
                oColunm.Editable = true;
                oColunm.Visible = true;
                oColunm.AffectsFormMode = true;
                oColunm.DataBind.SetBound(true, "", "SYS_62");
                oColunm.ValOn = "Y";
                oColunm.ValOff = "N";
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
                Business.escribirLog("crearFormulario: " + ex.Message);
            }
        }

        /// <summary>
        /// Creacion de items en formulario retenciones por articulo
        /// </summary>
        /// <param name="pVal">Tipo de evento</param>
        /// <param name="formUID">Identificador del formulario</param>
        private void CrearItemsRetSN(string FormUID)
        {
            try
            {
                oForm.DataSources.DBDataSources.Add("OWHT");
                oForm.DataSources.UserDataSources.Add("SYS_61", BoDataType.dt_LONG_NUMBER, 4);
                oForm.DataSources.UserDataSources.Add("SYS_62", BoDataType.dt_SHORT_TEXT, 1);
                oForm.DataSources.UserDataSources.Add("SYS_65", BoDataType.dt_LONG_NUMBER, 4);
                oForm.DataSources.UserDataSources.Add("SYS_66", BoDataType.dt_SHORT_TEXT, 10);
                oForm.DataSources.UserDataSources.Add("SYS_67", BoDataType.dt_SHORT_TEXT, 4);

                oItem = oForm.Items.Add("CardCode", BoFormItemTypes.it_EDIT);
                oItem.Left = 7;
                oItem.Top = 246;
                oItem.Height = 19;
                oItem.Width = 65;
                oItem.Visible = false;
                oEdit = (EditText)oItem.Specific;
                oEdit.Value = "";

                oItem = oForm.Items.Add("1", BoFormItemTypes.it_BUTTON);
                oItem.Left = 7;
                oItem.Top = 246;
                oItem.Height = 19;
                oItem.Width = 65;
                oButton = (Button)oItem.Specific;
                oButton.Caption = "OK";

                oItem = oForm.Items.Add("2", BoFormItemTypes.it_BUTTON);
                oItem.Left = 80;
                oItem.Top = 246;
                oItem.Height = 19;
                oItem.Width = 65;
                oButton = (Button)oItem.Specific;
                oButton.Caption = "Cancelar";

                oItem = oForm.Items.Add("3", BoFormItemTypes.it_MATRIX);
                oItem.Left = 5;
                oItem.Top = 5;
                oItem.Height = 240;
                oItem.Width = 320;
                oMatrix = (Matrix)oItem.Specific;
                oMatrix.SelectionMode = BoMatrixSelect.ms_Auto;

                oColunm = oMatrix.Columns.Add("0", BoFormItemTypes.it_EDIT);
                oColunm.TitleObject.Caption = "#";
                oColunm.Width = 22;
                oColunm.Editable = false;
                oColunm.Visible = true;
                oColunm.AffectsFormMode = true;
                oColunm.DataBind.SetBound(true, "", "SYS_61");

                oColunm = oMatrix.Columns.Add("3", BoFormItemTypes.it_EDIT);
                oColunm.TitleObject.Caption = "Código";
                oColunm.Width = 67;
                oColunm.Editable = false;
                oColunm.Visible = true;
                oColunm.AffectsFormMode = true;
                oColunm.DataBind.SetBound(true, "OWHT", "WTCode");

                oColunm = oMatrix.Columns.Add("1", BoFormItemTypes.it_EDIT);
                oColunm.TitleObject.Caption = "Descripción";
                oColunm.Width = 180;
                oColunm.Editable = false;
                oColunm.Visible = true;
                oColunm.AffectsFormMode = true;
                oColunm.DataBind.SetBound(true, "OWHT", "WTName");

                oColunm = oMatrix.Columns.Add("2", BoFormItemTypes.it_CHECK_BOX);
                oColunm.TitleObject.Caption = "Seleccionar";
                oColunm.Width = 33;
                oColunm.Editable = true;
                oColunm.Visible = true;
                oColunm.AffectsFormMode = true;
                oColunm.DataBind.SetBound(true, "", "SYS_62");
                oColunm.ValOn = "Y";
                oColunm.ValOff = "N";
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
                Business.escribirLog("crearFormulario: " + ex.Message);
            }
        }

        /// <summary>
        /// Creacion de items en formulario CIERRE FISCAL
        /// </summary>
        /// <param name="pVal">Tipo de evento</param>
        /// <param name="formUID">Identificador del formulario</param>
        private void CrearItemsCierre(string FormUID)
        {
            try
            {
                oForm.DataSources.UserDataSources.Add("dt_fIni", SAPbouiCOM.BoDataType.dt_DATE, 0);
                oForm.DataSources.UserDataSources.Add("dt_fFin", SAPbouiCOM.BoDataType.dt_DATE, 0);
                oForm.DataSources.UserDataSources.Add("EditToC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
                oForm.DataSources.UserDataSources.Add("EditFromC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
                oForm.DataSources.UserDataSources.Add("EditToS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
                oForm.DataSources.UserDataSources.Add("EditFromS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
                oForm.DataSources.DBDataSources.Add("OACT");
                oForm.DataSources.DBDataSources.Add("OCRD");
                oForm.DataSources.DBDataSources.Add("OTRC");

                oItem = oForm.Items.Add("Stc01", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 127;
                oItem.Top = 6;
                oItem.Height = 15;
                oItem.Width = 419;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                //oItem.TextStyle = (int)SAPbouiCOM.BoTextStyle.ts_UNDERLINE;
                oStatic = (SAPbouiCOM.StaticText)oItem.Specific;
                oStatic.Caption = "Parámetros generales";

                oItem = oForm.Items.Add("Stc02", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 127;
                oItem.Top = 21;
                oItem.Height = 15;
                oItem.Width = 419;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oItem.TextStyle = (int)SAPbouiCOM.BoTextStyle.ts_UNDERLINE;
                oStatic = (SAPbouiCOM.StaticText)oItem.Specific;
                oStatic.Caption = "Defina los parámetros generales para la ejecución del cierre fiscal";

                oItem = oForm.Items.Add("StcFecIni", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 12;
                oItem.Top = 87;
                oItem.Height = 14;
                oItem.Width = 161;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oItem.TextStyle = (int)SAPbouiCOM.BoTextStyle.ts_UNDERLINE;
                oStatic = (SAPbouiCOM.StaticText)oItem.Specific;
                oStatic.Caption = "Fecha Inicial";

                oItem = oForm.Items.Add("FechaIni", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem.Left = 174;
                oItem.Top = 87;
                oItem.Height = 14;
                oItem.Width = 132;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oItem.Enabled = true;
                oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                oEdit.DataBind.SetBound(true, "", "dt_fIni");

                oItem = oForm.Items.Add("StcFecFin", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 12;
                oItem.Top = 105;
                oItem.Height = 14;
                oItem.Width = 161;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oItem.TextStyle = (int)SAPbouiCOM.BoTextStyle.ts_UNDERLINE;
                oStatic = (SAPbouiCOM.StaticText)oItem.Specific;
                oStatic.Caption = "Fecha Final";

                oItem = oForm.Items.Add("FechaFin", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem.Left = 174;
                oItem.Top = 105;
                oItem.Height = 14;
                oItem.Width = 132;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oItem.Enabled = true;
                oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                oEdit.DataBind.SetBound(true, "", "dt_fFin");

                SAPbouiCOM.ChooseFromListCollection ocfls;
                ocfls = oForm.ChooseFromLists;
                SAPbouiCOM.ChooseFromList ocfl;
                SAPbouiCOM.ChooseFromListCreationParams cflcrepa;
                cflcrepa = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                cflcrepa.MultiSelection = false;
                cflcrepa.ObjectType = "1";
                cflcrepa.UniqueID = "CFL1";
                ocfl = ocfls.Add(cflcrepa);

                SAPbouiCOM.ChooseFromListCollection ocfls1;
                ocfls1 = oForm.ChooseFromLists;
                SAPbouiCOM.ChooseFromList ocfl1;
                SAPbouiCOM.ChooseFromListCreationParams cflcrepa1;
                cflcrepa1 = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                cflcrepa1.MultiSelection = false;
                cflcrepa1.ObjectType = "1";
                cflcrepa1.UniqueID = "CFL2";
                ocfl1 = ocfls1.Add(cflcrepa1);

                oItem = oForm.Items.Add("Stc03", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 12;
                oItem.Top = 142;
                oItem.Height = 14;
                oItem.Width = 122;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oItem.Visible = true;
                oItem.TextStyle = (int)SAPbouiCOM.BoTextStyle.ts_UNDERLINE;
                oStatic = (SAPbouiCOM.StaticText)oItem.Specific;
                oStatic.Caption = "Cuenta Inicial";

                oItem = oForm.Items.Add("ToAcct", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem.Left = 174;
                oItem.Top = 142;
                oItem.Height = 14;
                oItem.Width = 132;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oItem.Enabled = true;
                oItem.Visible = true;
                oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                oEdit.DataBind.SetBound(true, "", "EditToC");
                oEdit.ChooseFromListUID = "CFL1";
                oEdit.ChooseFromListAlias = "AcctCode";

                oItem = oForm.Items.Add("Stc04", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 12;
                oItem.Top = 160;
                oItem.Height = 14;
                oItem.Width = 122;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oItem.Visible = true;
                oItem.TextStyle = (int)SAPbouiCOM.BoTextStyle.ts_UNDERLINE;
                oStatic = (SAPbouiCOM.StaticText)oItem.Specific;
                oStatic.Caption = "Cuenta Final";

                oItem = oForm.Items.Add("FromAcct", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem.Left = 174;
                oItem.Top = 160;
                oItem.Height = 14;
                oItem.Width = 132;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oItem.Enabled = true;
                oItem.Visible = true;
                oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                oEdit.DataBind.SetBound(true, "", "EditFromC");
                oEdit.ChooseFromListUID = "CFL2";
                oEdit.ChooseFromListAlias = "AcctCode";

                SAPbouiCOM.ChooseFromListCollection ocfls2;
                ocfls2 = oForm.ChooseFromLists;
                SAPbouiCOM.ChooseFromList ocfl2;
                SAPbouiCOM.ChooseFromListCreationParams cflcrepa2;
                cflcrepa2 = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                cflcrepa2.MultiSelection = false;
                cflcrepa2.ObjectType = "2";
                cflcrepa2.UniqueID = "CFL3";
                ocfl2 = ocfls2.Add(cflcrepa2);

                SAPbouiCOM.ChooseFromListCollection ocfls3;
                ocfls3 = oForm.ChooseFromLists;
                SAPbouiCOM.ChooseFromList ocfl3;
                SAPbouiCOM.ChooseFromListCreationParams cflcrepa3;
                cflcrepa3 = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                cflcrepa3.MultiSelection = false;
                cflcrepa3.ObjectType = "2";
                cflcrepa3.UniqueID = "CFL4";
                ocfl3 = ocfls3.Add(cflcrepa3);

                //SAPbouiCOM.ChooseFromListCollection ocfls4;
                //ocfls4 = oForm.ChooseFromLists;
                //SAPbouiCOM.ChooseFromList ocfl4;
                //SAPbouiCOM.ChooseFromListCreationParams cflcrepa4;
                //cflcrepa4 = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                //cflcrepa4.MultiSelection = false;
                //cflcrepa4.ObjectType = "45";
                //cflcrepa4.UniqueID = "CFL5";
                //ocfl4 = ocfls4.Add(cflcrepa4);

                oItem = oForm.Items.Add("Stc05", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 12;
                oItem.Top = 197;
                oItem.Height = 14;
                oItem.Width = 122;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oItem.Visible = true;
                oItem.TextStyle = (int)SAPbouiCOM.BoTextStyle.ts_UNDERLINE;
                oStatic = (SAPbouiCOM.StaticText)oItem.Specific;
                oStatic.Caption = "Socio Inicial";

                oItem = oForm.Items.Add("ToSocio", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem.Left = 174;
                oItem.Top = 197;
                oItem.Height = 14;
                oItem.Width = 132;
                oItem.Enabled = true;
                oItem.Visible = true;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                oEdit.DataBind.SetBound(true, "", "EditToS");
                oEdit.ChooseFromListUID = "CFL3";
                oEdit.ChooseFromListAlias = "CardCode";

                oItem = oForm.Items.Add("Stc06", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 12;
                oItem.Top = 215;
                oItem.Height = 14;
                oItem.Width = 122;
                oItem.Visible = true;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oItem.TextStyle = (int)SAPbouiCOM.BoTextStyle.ts_UNDERLINE;
                oStatic = (SAPbouiCOM.StaticText)oItem.Specific;
                oStatic.Caption = "Socio Final";

                oItem = oForm.Items.Add("FromSocio", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem.Left = 174;
                oItem.Top = 215;
                oItem.Height = 14;
                oItem.Width = 132;
                oItem.Enabled = true;
                oItem.Visible = true;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                oEdit.DataBind.SetBound(true, "", "EditFromS");
                oEdit.ChooseFromListUID = "CFL4";
                oEdit.ChooseFromListAlias = "CardCode";

                oItem = oForm.Items.Add("Stc07", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 12;
                oItem.Top = 233;
                oItem.Height = 14;
                oItem.Width = 122;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oItem.Visible = true;
                oItem.TextStyle = (int)SAPbouiCOM.BoTextStyle.ts_UNDERLINE;
                oStatic = (SAPbouiCOM.StaticText)oItem.Specific;
                oStatic.Caption = "Codigo Transaccion";

                oItem = oForm.Items.Add("TransCode", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oItem.Left = 174;
                oItem.Top = 233;
                oItem.Height = 14;
                oItem.Width = 132;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oComboBox = oItem.Specific;
                SAPbobsCOM.Recordset oRecPro = null;
                oRecPro = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                oRecPro.DoQuery(string.Format(Consultas.Default.ListTransCode));
                for (int i = 0; i <= oComboBox.ValidValues.Count - 1; i++)
                {
                    oComboBox.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }
                oComboBox.ValidValues.Add("", "");
                while (oRecPro.EoF == false)
                {
                    oComboBox.ValidValues.Add(oRecPro.Fields.Item(0).Value, oRecPro.Fields.Item(1).Value);
                    oRecPro.MoveNext();
                }

                oItem = oForm.Items.Add("Stc08", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 12;
                oItem.Top = 251;
                oItem.Height = 14;
                oItem.Width = 122;
                oItem.Visible = true;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oItem.TextStyle = (int)SAPbouiCOM.BoTextStyle.ts_UNDERLINE;
                oStatic = (SAPbouiCOM.StaticText)oItem.Specific;
                oStatic.Caption = "Año del cierre fiscal";

                oItem = oForm.Items.Add("Year", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oItem.Left = 174;
                oItem.Top = 251;
                oItem.Height = 14;
                oItem.Width = 132;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oComboBox = oItem.Specific;
                oRecPro = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                oRecPro.DoQuery(string.Format(Consultas.Default.Years));
                for (int i = 0; i <= oComboBox.ValidValues.Count - 1; i++)
                {
                    oComboBox.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }
                oComboBox.ValidValues.Add("", "");
                while (oRecPro.EoF == false)
                {
                    oComboBox.ValidValues.Add(oRecPro.Fields.Item(0).Value, oRecPro.Fields.Item(0).Value);
                    oRecPro.MoveNext();
                }

                oForm.DataSources.DataTables.Add("DTCIERRE");
                oItem = oForm.Items.Add("Grid", SAPbouiCOM.BoFormItemTypes.it_GRID);
                oItem.Left = 30;
                oItem.Top = 65;
                oItem.Height = 270;
                oItem.Width = 555;
                oItem.FromPane = 2;
                oItem.ToPane = 2;
                oItem.Visible = false;
                oGrid = (SAPbouiCOM.Grid)oItem.Specific;
                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None;

                oItem = oForm.Items.Add("2", BoFormItemTypes.it_BUTTON);
                oItem.Left = 371;
                oItem.Top = 488;
                oItem.Height = 19;
                oItem.Width = 65;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Type = SAPbouiCOM.BoButtonTypes.bt_Caption;
                oButton.Caption = "Cancelar";

                oItem = oForm.Items.Add("3", BoFormItemTypes.it_BUTTON);
                oItem.Left = 441;
                oItem.Top = 488;
                oItem.Height = 19;
                oItem.Width = 65;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Type = SAPbouiCOM.BoButtonTypes.bt_Caption;
                oButton.Caption = "< Atr&ás";

                oItem = oForm.Items.Add("4", BoFormItemTypes.it_BUTTON);
                oItem.Left = 511;
                oItem.Top = 488;
                oItem.Height = 19;
                oItem.Width = 65;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Type = SAPbouiCOM.BoButtonTypes.bt_Caption;
                oButton.Caption = "Sig&uiente >";

                oItem = oForm.Items.Add("129", BoFormItemTypes.it_BUTTON);
                oItem.Left = 511;
                oItem.Top = 488;
                oItem.Height = 19;
                oItem.Width = 65;
                oItem.FromPane = 2;
                oItem.ToPane = 2;
                oItem.Visible = false;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Type = SAPbouiCOM.BoButtonTypes.bt_Caption;
                oButton.Caption = "Finali&zar";
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
                Business.escribirLog("crearFormulario: " + ex.Message);
            }
        }


        /// <summary>
        /// Manaejo de eventos formulario configuracion FE en Parametrizaciones Generales
        /// </summary>
        /// <param name="pVal">Tipo de evento</param>
        /// <param name="formUID">Identificador del formulario</param>
        private void EventFieldVentas(ItemEvent pVal, string FormUID)
        {
            if ((pVal.FormType == 133) && ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && (pVal.Before_Action == true)) || (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && (pVal.Before_Action == false))))
            //  if ((pVal.FormType == 133 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true) || (pVal.FormType == 133 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && pVal.Before_Action == true))
            {
                try
                {
                    if (SBO_Application.Menus.Exists("ReclasificacionOPCH")) SBO_Application.Menus.RemoveEx("ReclasificacionOPCH");
                    if (SBO_Application.Menus.Exists("ReclasificacionOINV") == false)
                    {
                        SAPbouiCOM.MenuItem oMenuItem = null;
                        SAPbouiCOM.Menus oMenus = null;
                        SAPbouiCOM.MenuCreationParams oCreationPackage = null;
                        oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        oCreationPackage.UniqueID = "ReclasificacionOINV";
                        oCreationPackage.String = "Reclasificación de retenciones";
                        //Agregado 09/16/2019                        
                        string[] paths = { @"" + rutaDocs + "", "Logs", "Iconos", "Reclasificacion" };
                        string fullPath = Path.Combine(paths);
                        oCreationPackage.Image = fullPath + ".bmp";
                        //---------------------------------------------------------
                        oCreationPackage.Enabled = true;
                        oMenuItem = SBO_Application.Menus.Item("1280"); // Data'1280
                        oMenus = oMenuItem.SubMenus;
                        oMenus.AddEx(oCreationPackage);
                        GC.Collect();
                    }
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox("Item Event \n" + ex.Message);
                }
            }
            if ((pVal.FormType == 133) && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE || pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE))// && (pVal.Before_Action == true))
            {
                try
                {
                    if (SBO_Application.Menus.Exists("ReclasificacionOINV")) SBO_Application.Menus.RemoveEx("ReclasificacionOINV");
                    GC.Collect();
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                }
            }
            /* if (((pVal.ItemUID == "38") && (pVal.ColUID == "1") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) && (pVal.Before_Action == false)))
             {
                 string codigoArti = "";
                 string WtLiable;

                 try
                 {
                     oForm = SBO_Application.Forms.Item(FormUID);

                     oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                     oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific;
                     codigoArti = oEdit.Value;

                     SAPbobsCOM.Items item = (SAPbobsCOM.Items)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                     item.GetByKey(codigoArti);
                     WtLiable = item.UserFields.Fields.Item("U_SCL_WTLiable").Value;

                     if (WtLiable == "Y")
                     {
                         string sSQL = "";
                         Recordset oRS;
                         oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                         sSQL = string.Format(Consultas.Default.CulculoRet, codigoArti);
                         escribirLog("ConsultaRet: " + sSQL);
                         oRS.DoQuery(sSQL);
                         if (oRS.RecordCount > 0)
                         {
                             string CodRet = "";
                             string Cantidad = "";
                             string valor = "";
                             double prctRet = 0;

                             CodRet = oRS.Fields.Item("WTCode").Value;

                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific;
                             Cantidad = oEdit.Value;

                             SAPbobsCOM.WithholdingTaxCodes oWithholding;
                             oWithholding = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
                             oWithholding.GetByKey(CodRet);
                             prctRet = oWithholding.BaseAmount;


                             oEdit = oMatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific;
                             valor = oEdit.Value;
                             valor = Regex.Replace(valor, "[$|COP|USD|EUR]", "");
                             valor = valor.Replace(DatosGlobalesFE.sapMillar, "");

                             if (String.IsNullOrEmpty(valor))
                             {
                                 valor = "0";
                             }

                             if (String.IsNullOrEmpty(Cantidad))
                             {
                                 Cantidad = "0";
                             }

                             if (!valor.Contains(DatosGlobalesFE.localdecimal) || !Cantidad.Contains(DatosGlobalesFE.localdecimal))
                             {
                                 switch (DatosGlobalesFE.localdecimal)
                                 {
                                     case ("."):
                                         valor = valor.Replace(",", ".");
                                         Cantidad = Cantidad.Replace(",", ".");
                                         break;
                                     case (","):
                                         valor = valor.Replace(".", ",");
                                         Cantidad = Cantidad.Replace(".", ",");
                                         break;
                                 }
                             }

                             double valorRet = 0;
                             valorRet = (Convert.ToDouble(valor) * Convert.ToDouble(Cantidad)) * (prctRet / 100);

                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Cod_Ret").Cells.Item(pVal.Row).Specific;
                             oEdit.Value = CodRet;

                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Ret_Val").Cells.Item(pVal.Row).Specific;
                             oEdit.Value = valorRet.ToString();

                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Ret_Prct").Cells.Item(pVal.Row).Specific;
                             oEdit.Value = prctRet.ToString();
                         }
                         System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                         oRS = null;
                         GC.Collect();
                     }
                     System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                     item = null;
                     GC.Collect();
                 }
                 catch (Exception ex)
                 {
                     escribirLog("AddLine: " + ex.Message);
                     SBO_Application.MessageBox(ex.Message);
                 }
             }

             if (((pVal.ItemUID == "38") && (pVal.ColUID == "11") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) && (pVal.Before_Action == false)))
             {
                 string codigoArti = "";
                 string WtLiable;
                 try
                 {
                     oForm = SBO_Application.Forms.Item(FormUID);

                     oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                     oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific;
                     codigoArti = oEdit.Value;

                     SAPbobsCOM.Items item = (SAPbobsCOM.Items)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                     item.GetByKey(codigoArti);
                     WtLiable = item.UserFields.Fields.Item("U_SCL_WTLiable").Value;

                     if (WtLiable == "Y")
                     {
                         string sSQL = "";
                         Recordset oRS;
                         oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                         sSQL = string.Format(Consultas.Default.CulculoRet, codigoArti);
                         escribirLog("ConsultaRet: " + sSQL);
                         oRS.DoQuery(sSQL);
                         if (oRS.RecordCount > 0)
                         {
                             string CodRet = "";
                             string Cantidad = "";
                             string valor = "";
                             double prctRet = 0;

                             CodRet = oRS.Fields.Item("WTCode").Value;

                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific;
                             Cantidad = oEdit.Value;

                             SAPbobsCOM.WithholdingTaxCodes oWithholding;
                             oWithholding = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
                             oWithholding.GetByKey(CodRet);
                             prctRet = oWithholding.BaseAmount;


                             oEdit = oMatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific;
                             valor = oEdit.Value;
                             valor = Regex.Replace(valor, "[$|COP|USD|EUR]", "");
                             valor = valor.Replace(DatosGlobalesFE.sapMillar, "");

                             if (String.IsNullOrEmpty(valor))
                             {
                                 valor = "0";
                             }

                             if (String.IsNullOrEmpty(Cantidad))
                             {
                                 Cantidad = "0";
                             }

                             if (!valor.Contains(DatosGlobalesFE.localdecimal) || !Cantidad.Contains(DatosGlobalesFE.localdecimal))
                             {
                                 switch (DatosGlobalesFE.localdecimal)
                                 {
                                     case ("."):
                                         valor = valor.Replace(",", ".");
                                         Cantidad = Cantidad.Replace(",", ".");
                                         break;
                                     case (","):
                                         valor = valor.Replace(".", ",");
                                         Cantidad = Cantidad.Replace(".", ",");
                                         break;
                                 }
                             }

                             double valorRet = 0;
                             valorRet = (Convert.ToDouble(valor) * Convert.ToDouble(Cantidad)) * (prctRet / 100);

                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Cod_Ret").Cells.Item(pVal.Row).Specific;
                             oEdit.Value = CodRet;

                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Ret_Val").Cells.Item(pVal.Row).Specific;
                             oEdit.Value = valorRet.ToString();

                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Ret_Prct").Cells.Item(pVal.Row).Specific;
                             oEdit.Value = prctRet.ToString();
                         }
                         System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                         oRS = null;
                         GC.Collect();
                     }
                     System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                     item = null;
                     GC.Collect();
                 }
                 catch (Exception ex)
                 {
                     escribirLog("ChangeQuantity: " + ex.Message);
                     SBO_Application.MessageBox(ex.Message);
                 }
             }*/



            //verificacion para el campo de usuario retenciones por articulo
            if (((pVal.ItemUID == "38") && (pVal.ColUID == "U_SCL_Cod_Ret") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) && (pVal.Before_Action == false)))
            {
                string RetArt = "";
                string codigoArti = "";

                try
                {
                    string sSQL = "";
                    oForm = SBO_Application.Forms.Item(FormUID);
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                    oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific;
                    codigoArti = oEdit.Value;
                    oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Cod_Ret").Cells.Item(pVal.Row).Specific;
                    RetArt = oEdit.Value;

                    Recordset oRS;
                    oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    sSQL = string.Format(Consultas.Default.CalculoRetItem, codigoArti, RetArt);
                    escribirLog("ConsultaRet: " + RetArt);// escribirLog("ConsultaRet: " + sSQL);
                    oRS.DoQuery(sSQL);
                    if (oRS.RecordCount > 0)
                    {
                        string CodRet = "";
                        string Cantidad = "";
                        string valor = "";
                        double prctRet = 0;

                        CodRet = oRS.Fields.Item("WTCode").Value;

                        oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific;
                        Cantidad = oEdit.Value;

                        SAPbobsCOM.WithholdingTaxCodes oWithholding;
                        oWithholding = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
                        oWithholding.GetByKey(CodRet);
                        prctRet = oWithholding.BaseAmount;


                        oEdit = oMatrix.Columns.Item("21").Cells.Item(pVal.Row).Specific;
                        valor = oEdit.Value;
                        valor = Regex.Replace(valor, "[$|COP|USD|EUR]", "");
                        valor = valor.Replace(DatosGlobalesFE.sapMillar, "");

                        if (String.IsNullOrEmpty(valor))
                        {
                            valor = "0";
                        }

                        if (String.IsNullOrEmpty(Cantidad))
                        {
                            Cantidad = "0";
                        }

                        if (!valor.Contains(DatosGlobalesFE.localdecimal) || !Cantidad.Contains(DatosGlobalesFE.localdecimal))
                        {
                            switch (DatosGlobalesFE.localdecimal)
                            {
                                case ("."):
                                    valor = valor.Replace(",", ".");
                                    Cantidad = Cantidad.Replace(",", ".");
                                    break;
                                case (","):
                                    valor = valor.Replace(".", ",");
                                    Cantidad = Cantidad.Replace(".", ",");
                                    break;
                            }
                        }

                        double valorRet = 0;
                        valorRet = (Convert.ToDouble(valor) * Convert.ToDouble(Cantidad)) * (prctRet / 100);

                        oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Ret_Val").Cells.Item(pVal.Row).Specific;
                        oEdit.Value = valorRet.ToString();

                        oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Ret_Prct").Cells.Item(pVal.Row).Specific;
                        oEdit.Value = prctRet.ToString();
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                    oRS = null;
                    GC.Collect();
                }
                catch (Exception ex)
                {
                    escribirLog("ChangePrice: " + ex.Message);
                    SBO_Application.MessageBox(ex.Message);
                }
            }

            if (((pVal.ItemUID == "38") && (pVal.ColUID == "14") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) && (pVal.Before_Action == false)))

            {
                string codigoArti = "";
                string WtLiable;

                try
                {
                    oForm = SBO_Application.Forms.Item(FormUID);

                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                    oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific;
                    codigoArti = oEdit.Value;

                    SAPbobsCOM.Items item = (SAPbobsCOM.Items)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                    item.GetByKey(codigoArti);
                    WtLiable = item.UserFields.Fields.Item("U_SCL_WTLiable").Value;

                    if (WtLiable == "Y")
                    {
                        string sSQL = "";
                        Recordset oRS;
                        oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        sSQL = string.Format(Consultas.Default.CalItemRet, codigoArti);
                        oRS.DoQuery(sSQL);
                        escribirLog("ConsultaRet: " + sSQL);

                        if (oRS.RecordCount > 0)
                        {
                            string CodRet = "";
                            string Cantidad = "";
                            string valor = "";
                            double prctRet = 0;

                            CodRet = oRS.Fields.Item("WTCode").Value;

                            oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific;
                            Cantidad = oEdit.Value;

                            SAPbobsCOM.WithholdingTaxCodes oWithholding;
                            oWithholding = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
                            oWithholding.GetByKey(CodRet);
                            prctRet = oWithholding.BaseAmount;


                            oEdit = oMatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific;
                            valor = oEdit.Value;
                            valor = Regex.Replace(valor, "[$|COP|USD|EUR]", "");
                            valor = valor.Replace(DatosGlobalesFE.sapMillar, "");

                            if (String.IsNullOrEmpty(valor))
                            {
                                valor = "0";
                            }

                            if (String.IsNullOrEmpty(Cantidad))
                            {
                                Cantidad = "0";
                            }

                            if (!valor.Contains(DatosGlobalesFE.localdecimal) || !Cantidad.Contains(DatosGlobalesFE.localdecimal))
                            {
                                switch (DatosGlobalesFE.localdecimal)
                                {
                                    case ("."):
                                        valor = valor.Replace(",", ".");
                                        Cantidad = Cantidad.Replace(",", ".");
                                        break;
                                    case (","):
                                        valor = valor.Replace(".", ",");
                                        Cantidad = Cantidad.Replace(".", ",");
                                        break;
                                }
                            }

                            double valorRet = 0;
                            valorRet = (Convert.ToDouble(valor) * Convert.ToDouble(Cantidad)) * (prctRet / 100);

                            oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Cod_Ret").Cells.Item(pVal.Row).Specific;
                            oEdit.Value = CodRet;

                            oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Ret_Val").Cells.Item(pVal.Row).Specific;
                            oEdit.Value = valorRet.ToString();

                            oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Ret_Prct").Cells.Item(pVal.Row).Specific;
                            oEdit.Value = prctRet.ToString();
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                        oRS = null;
                        GC.Collect();
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                    item = null;
                    GC.Collect();
                }
                catch (Exception ex)
                {
                    escribirLog("ChangePrice: " + ex.Message);
                    SBO_Application.MessageBox(ex.Message);
                }
            }

            /* if (pVal.ItemUID == "1" && pVal.FormMode == 3 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == true)
             {
                 try
                 {
                     string CardCode = "";
                     string TotalDoc = "";
                     oForm = SBO_Application.Forms.Item(FormUID);

                     oEdit = oForm.Items.Item("4").Specific;
                     CardCode = oEdit.Value;

                     oEdit = oForm.Items.Item("22").Specific;

                     TotalDoc = oEdit.Value;

                     TotalDoc = oEdit.Value;
                     TotalDoc = Regex.Replace(TotalDoc, "[$|COP|USD|EUR]", "");
                     TotalDoc = TotalDoc.Replace(DatosGlobalesFE.sapMillar, "");
                     if (string.IsNullOrEmpty(TotalDoc))
                     {
                         TotalDoc = "0";
                     }

                     if (!TotalDoc.Contains(DatosGlobalesFE.localdecimal))
                     {
                         switch (DatosGlobalesFE.localdecimal)
                         {
                             case ("."):
                                 TotalDoc = TotalDoc.Replace(",", ".");
                                 break;
                             case (","):
                                 TotalDoc = TotalDoc.Replace(".", ",");
                                 break;
                         }
                     }

                     oItem = oForm.Items.Item("91");
                     oItem.Click();

                     oForm = SBO_Application.Forms.ActiveForm;
                     oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;

                     int row = 1;
                     while (row <= oMatrix.RowCount)
                     {
                         int codGasto;
                         string strGasto;
                         oEdit = oMatrix.Columns.Item("1").Cells.Item(row).Specific;
                         codGasto = Convert.ToInt32(oEdit.Value);

                         SAPbobsCOM.AdditionalExpenses oExpenses;
                         oExpenses = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAdditionalExpenses);
                         oExpenses.GetByKey(codGasto);
                         strGasto = Convert.ToString(oExpenses.Name);

                         Recordset oRS;
                         string sSQL = "";
                         oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                         sSQL = string.Format(Consultas.Default.CalcuAutoRet, CardCode, strGasto);
                         escribirLog("ConsultaAutoRet: " + sSQL);
                         oRS.DoQuery(sSQL);

                         if (oRS.RecordCount > 0)
                         {
                             double prctRet = 0;
                             string CodRet = "";

                             CodRet = oRS.Fields.Item("WTCode").Value;

                             SAPbobsCOM.WithholdingTaxCodes oWithholding;
                             oWithholding = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
                             oWithholding.GetByKey(CodRet);
                             prctRet = oWithholding.BaseAmount;

                             double valorret = 0;
                             valorret = Convert.ToDouble(TotalDoc) * (prctRet / 100);

                             if (TotalDoc.Contains(DatosGlobalesFE.localdecimal))
                             {
                                 oEdit = oMatrix.Columns.Item("3").Cells.Item(row).Specific;
                                 oEdit.Value = valorret.ToString().Replace(DatosGlobalesFE.localdecimal, DatosGlobalesFE.sapdecimal);
                             }
                             else
                             {
                                 oEdit = oMatrix.Columns.Item("3").Cells.Item(row).Specific;
                                 oEdit.Value = valorret.ToString();
                             }

                             System.Runtime.InteropServices.Marshal.ReleaseComObject(oWithholding);
                             oWithholding = null;
                             GC.Collect();
                         }

                         row++;

                         System.Runtime.InteropServices.Marshal.ReleaseComObject(oExpenses);
                         oExpenses = null;
                         GC.Collect();
                     }

                     if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                     {
                         oItem = oForm.Items.Item("1");
                         oItem.Click();
                     }
                     if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                     {
                         oItem = oForm.Items.Item("1");
                         oItem.Click();
                     }

                     oForm = SBO_Application.Forms.Item(FormUID);
                     oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;

                     System.Data.DataTable dt = new System.Data.DataTable();
                     dt.Columns.Add(new System.Data.DataColumn("1", typeof(string)));
                     dt.Columns.Add(new System.Data.DataColumn("2", typeof(double)));
                     dt.Columns.Add(new System.Data.DataColumn("3", typeof(double)));

                     for (int i = 1; i <= oMatrix.RowCount; i++)
                     {
                         string codRet = "";
                         string Valor = "";
                         string price = "";
                         string cantidad = "";

                         oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Cod_Ret").Cells.Item(i).Specific;
                         codRet = oEdit.Value;

                         if (!string.IsNullOrEmpty(codRet))
                         {
                             DataRow newrow = dt.NewRow();
                             newrow[0] = codRet;

                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Ret_Val").Cells.Item(i).Specific;
                             Valor = oEdit.Value;

                             oEdit = oMatrix.Columns.Item("14").Cells.Item(i).Specific;
                             price = oEdit.Value;
                             price = Regex.Replace(price, "[$|COP|USD|EUR]", "");
                             price = price.Replace(DatosGlobalesFE.sapMillar, "");

                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(i).Specific;
                             cantidad = oEdit.Value;

                             if (String.IsNullOrEmpty(price))
                             {
                                 price = "0";
                             }

                             if (String.IsNullOrEmpty(cantidad))
                             {
                                 cantidad = "0";
                             }

                             if (!Valor.Contains(DatosGlobalesFE.localdecimal) || !price.Contains(DatosGlobalesFE.localdecimal) || !cantidad.Contains(DatosGlobalesFE.localdecimal))
                             {
                                 switch (DatosGlobalesFE.localdecimal)
                                 {
                                     case ("."):
                                         Valor = Valor.Replace(",", ".");
                                         price = price.Replace(",", ".");
                                         cantidad = cantidad.Replace(",", ".");
                                         break;
                                     case (","):
                                         Valor = Valor.Replace(".", ",");
                                         price = price.Replace(".", ",");
                                         cantidad = cantidad.Replace(".", ",");
                                         break;
                                 }
                             }
                             double valorBase = 0;
                             valorBase = Convert.ToDouble(price) * Convert.ToDouble(cantidad);
                             escribirLog("U_SCL_Ret_Val: " + Valor);
                             escribirLog("valorbase: " + valorBase);
                             newrow[1] = Valor;
                             newrow[2] = valorBase;
                             dt.Rows.Add(newrow);
                         }
                     }

                     if (dt != null && dt.Rows.Count > 0)
                     {
                         SBO_Application.Menus.Item("5897").Activate();
                         var query = from r in dt.AsEnumerable()
                                     group r by r.Field<string>(0) into groupedTable
                                     select new
                                     {
                                         id = groupedTable.Key,
                                         sumOfValue = groupedTable.Sum(s => s.Field<double>("2")),
                                         valueBase = groupedTable.Sum(s => s.Field<double>("3"))
                                     };

                         System.Data.DataTable newDt = ConvertToDataTable(query);

                         oForm = SBO_Application.Forms.ActiveForm;
                         oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("6").Specific;

                         oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("68").Cells.Item(1).Specific;
                         string docBase = oEdit.Value;

                         if(docBase == "-1")
                         {
                             int rowmatrix = 1;
                             for (int i = 0; i < newDt.Rows.Count;)
                             {
                                 bool existRet = false;
                                 SAPbobsCOM.WithholdingTaxCodes oWithholding;
                                 oWithholding = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
                                 existRet = oWithholding.GetByKey(newDt.Rows[i]["id"].ToString());

                                 if (!string.IsNullOrEmpty(newDt.Rows[i]["id"].ToString()) && existRet)
                                 {
                                     oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(rowmatrix).Specific;
                                     escribirLog("codRet: " + newDt.Rows[i]["id"].ToString());
                                     oEdit.Value = newDt.Rows[i]["id"].ToString();

                                     oColunm = oMatrix.Columns.Item("14");
                                     if (oColunm.Editable == true)
                                     {
                                         oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("14").Cells.Item(rowmatrix).Specific;
                                         escribirLog("SumValorRet: " + newDt.Rows[i]["sumOfValue"].ToString());
                                         oEdit.Value = newDt.Rows[i]["sumOfValue"].ToString();
                                     }
                                     else
                                     {
                                         oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("28").Cells.Item(rowmatrix).Specific;
                                         escribirLog("SumValorRet: " + newDt.Rows[i]["sumOfValue"].ToString());
                                         oEdit.Value = newDt.Rows[i]["sumOfValue"].ToString();
                                     }

                                     oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(rowmatrix).Specific;
                                     if (string.IsNullOrEmpty(oEdit.Value))
                                     {
                                         string msg = "";
                                         msg = "Retencion " + newDt.Rows[i]["id"].ToString() + " no asignada al socio de negocio";
                                         SBO_Application.StatusBar.SetSystemMessage(msg, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
                                     }
                                     else
                                     {
                                         rowmatrix++;
                                     }
                                 }
                                 if (!string.IsNullOrEmpty(newDt.Rows[i]["id"].ToString()) && existRet == false)
                                 {
                                     string msg = "";
                                     msg = "La Retencion " + newDt.Rows[i]["id"].ToString() + " no existe";
                                     SBO_Application.StatusBar.SetSystemMessage(msg, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
                                 }

                                 i++;

                                 System.Runtime.InteropServices.Marshal.ReleaseComObject(oWithholding);
                                 oWithholding = null;
                                 GC.Collect();
                             }
                         }
                         else
                         {
                             for (int i = 1; i <= oMatrix.RowCount; i++)
                             {
                                 bool existRet = false;
                                 string retActual = "";
                                 oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific;
                                 retActual = oEdit.Value;
                                 for (int j = 0; j < newDt.Rows.Count; j++)
                                 {
                                     string retNueva = "";
                                     retNueva = newDt.Rows[j]["id"].ToString();
                                     if(retActual == retNueva)
                                     {
                                         existRet = true;
                                         oColunm = oMatrix.Columns.Item("14");
                                         if (oColunm.Editable == true)
                                         {
                                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("14").Cells.Item(i).Specific;
                                             escribirLog("SumValorRet: " + newDt.Rows[j]["sumOfValue"].ToString());
                                             oEdit.Value = newDt.Rows[j]["sumOfValue"].ToString();
                                         }
                                         else
                                         {
                                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("28").Cells.Item(i).Specific;
                                             escribirLog("SumValorRet: " + newDt.Rows[j]["sumOfValue"].ToString());
                                             oEdit.Value = newDt.Rows[j]["sumOfValue"].ToString();
                                         }
                                     }
                                 }
                                 if(existRet == false)
                                 {
                                     oMatrix.DeleteRow(i);
                                 }
                             }
                         }

                         if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                         {
                             oItem = oForm.Items.Item("1");
                             oItem.Click();
                         }
                         if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                         {
                             oItem = oForm.Items.Item("1");
                             oItem.Click();
                         }
                     }
                 }
                 catch (Exception ex)
                 {
                     escribirLog("CalcuRet: " + ex.Message);
                     SBO_Application.MessageBox(ex.Message);
                 }
             }*/
        }

        /// <summary>
        /// Manaejo de eventos formulario configuracion FE en Parametrizaciones Generales
        /// </summary>
        /// <param name="pVal">Tipo de evento</param>
        /// <param name="formUID">Identificador del formulario</param>
        private void EventFieldCompras(ItemEvent pVal, string FormUID)
        {
            if ((pVal.FormType == 141) && ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && (pVal.Before_Action == true)) || (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && (pVal.Before_Action == false))))
            {
                try
                {
                    if (SBO_Application.Menus.Exists("ReclasificacionOINV")) SBO_Application.Menus.RemoveEx("ReclasificacionOINV");
                    if (SBO_Application.Menus.Exists("ReclasificacionOPCH") == false)
                    {
                        SAPbouiCOM.MenuItem oMenuItem = null;
                        SAPbouiCOM.Menus oMenus = null;
                        SAPbouiCOM.MenuCreationParams oCreationPackage = null;
                        oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        oCreationPackage.UniqueID = "ReclasificacionOPCH";
                        oCreationPackage.String = "Reclasificación de retenciones";
                        // oCreationPackage.Image = @"C:\Rec2.bmp";
                        SAPbobsCOM.Recordset oRecordset;
                        oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string query = String.Format(Properties.Resources.IPServidor);
                        oRecordset.DoQuery(query);
                        string ip = oRecordset.Fields.Item("U_SCL_RutaInf").Value;
                        string[] paths = { @"\\" + ip + "", "Addon SCL Colombia", "Iconos", "Reclasificacion" };
                        string fullPath = Path.Combine(paths);
                        oCreationPackage.Image = fullPath + ".bmp";
                        oCreationPackage.Enabled = true;
                        oMenuItem = SBO_Application.Menus.Item("1280"); // Data'1280
                        oMenus = oMenuItem.SubMenus;
                        oMenus.AddEx(oCreationPackage);
                        GC.Collect();
                    }
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                }
            }

            if ((pVal.FormType == 141) && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE || pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE))// && (pVal.Before_Action == true))
            {
                try
                {
                    if (SBO_Application.Menus.Exists("ReclasificacionOPCH")) SBO_Application.Menus.RemoveEx("ReclasificacionOPCH");
                    GC.Collect();
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                }
            }

            /*
            if ((pVal.FormType == 143) && ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && (pVal.Before_Action == false)) && (pVal.ActionSuccess == true) && (pVal.ItemUID == "1") && pVal.FormMode == 3))
            {
                var items = new ArrayList();
                string cntaIVAMV, undMed, SN;
                int num;
                int contRI = 0;
                int contAS = 0;
                IVA_Mayor.ArticuloIVA art;

                oForm = SBO_Application.Forms.Item(FormUID);
                oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("8").Specific;
                int docNum = Convert.ToInt32(oEdit.Value.ToString()) - 1;///KKKKKKKKKKKKKKKKK
                string sSQL = "";
                Recordset oRecordSet;
                oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                sSQL = string.Format(Consultas.Default.CntaIVAMV);
                oRecordSet.DoQuery(sSQL);
                if (oRecordSet.RecordCount > 0)
                {
                    cntaIVAMV = Convert.ToString(oRecordSet.Fields.Item(0).Value);
                    double tasa = Convert.ToDouble(oRecordSet.Fields.Item(1).Value);
                    string cntImp = Convert.ToString(oRecordSet.Fields.Item(2).Value);
                    sSQL = string.Format(Consultas.Default.CodIVAMV, docNum);
                    oRecordSet.DoQuery(sSQL);
                    
                    if (oRecordSet.RecordCount > 0)
                    {                         
                        undMed = Convert.ToString(oRecordSet.Fields.Item(5).Value);
                        num = Convert.ToInt32(oRecordSet.Fields.Item(7).Value);
                        SN = Convert.ToString(oRecordSet.Fields.Item(8).Value);
                        string invent = "";
                        double total = 0;
                        double costoActual = 0;
                        while (oRecordSet.EoF == false)
                        {
                            invent = Convert.ToString(oRecordSet.Fields.Item(9).Value);
                            if (invent.Equals("Y"))
                            {
                                contRI++;
                                art = new IVA_Mayor.ArticuloIVA();
                                art.codigo = Convert.ToString(oRecordSet.Fields.Item(0).Value);
                                art.cantidad = Convert.ToInt32(oRecordSet.Fields.Item(1).Value);
                                art.almacen = Convert.ToString(oRecordSet.Fields.Item(6).Value);
                                //items.Add(Convert.ToString(oRecordSet.Fields.Item(0).Value));
                                int cantidad = Convert.ToInt32(oRecordSet.Fields.Item(1).Value);
                                double precio = Convert.ToDouble(oRecordSet.Fields.Item(2).Value);
                                //string almacen = Convert.ToString(oRecordSet.Fields.Item(6).Value);
                                double precioAct = ((precio * tasa) / 100);
                                precioAct = precioAct * cantidad;
                                // double precioAct = ((precio * tasa) / 100)  ;
                                costoActual = precio * cantidad;
                                art.costoActual = costoActual;
                                art.costoNuevo = precioAct;
                                items.Add(art);
                            }else{
                                contAS++;
                                int cantidad = Convert.ToInt32(oRecordSet.Fields.Item(1).Value);
                                double precio = Convert.ToDouble(oRecordSet.Fields.Item(2).Value);
                                double precioAct = ((precio * tasa) / 100);
                                total += precioAct * cantidad;
                            }
                            oRecordSet.MoveNext();
                        }
                        //oConnection.ConCompany(oCompany, SBO_Application);
                        string sessionID = oConnection.ConexionServiceLayer();
                        if (contRI > 0)
                        {
                            addMaterialRevaluationIVAMV(items, cntaIVAMV, SN, num, sessionID);//lllllll
                            CreateJournalEntryIVAMV(cntImp, costoActual, cntaIVAMV, SN, num, sessionID, 2);
                        }
                        if (contAS > 0) CreateJournalEntryIVAMV(cntaIVAMV, total, cntImp, SN, num, sessionID, 1);
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    return;
                }
            }
            */
            /*if (((pVal.ItemUID == "38") && (pVal.ColUID == "1") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) && (pVal.Before_Action == false)))
             {
                 string codigoArti = "";
                 string WtLiable;

                 try
                 {
                     oForm = SBO_Application.Forms.Item(FormUID);

                     oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                     oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific;
                     codigoArti = oEdit.Value;

                     SAPbobsCOM.Items item = (SAPbobsCOM.Items)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                     item.GetByKey(codigoArti);
                     WtLiable = item.UserFields.Fields.Item("U_SCL_WTLiable").Value;

                     if (WtLiable == "Y")
                     {
                         string sSQL = "";
                         Recordset oRS;
                         oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                         sSQL = string.Format(Consultas.Default.CulculoRet, codigoArti);
                         escribirLog("ConsultaRet: " + sSQL);
                         oRS.DoQuery(sSQL);
                         if (oRS.RecordCount > 0)
                         {
                             string CodRet = "";
                             string Cantidad = "";
                             string valor = "";
                             double prctRet = 0;

                             CodRet = oRS.Fields.Item("WTCode").Value;

                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific;
                             Cantidad = oEdit.Value;

                             SAPbobsCOM.WithholdingTaxCodes oWithholding;
                             oWithholding = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
                             oWithholding.GetByKey(CodRet);
                             prctRet = oWithholding.BaseAmount;


                             oEdit = oMatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific;
                             valor = oEdit.Value;
                             valor = Regex.Replace(valor, "[$|COP|USD|EUR]", "");
                             valor = valor.Replace(DatosGlobalesFE.sapMillar, "");

                             if (String.IsNullOrEmpty(valor))
                             {
                                 valor = "0";
                             }

                             if (String.IsNullOrEmpty(Cantidad))
                             {
                                 Cantidad = "0";
                             }

                             if (!valor.Contains(DatosGlobalesFE.localdecimal) || !Cantidad.Contains(DatosGlobalesFE.localdecimal))
                             {
                                 switch (DatosGlobalesFE.localdecimal)
                                 {
                                     case ("."):
                                         valor = valor.Replace(",", ".");
                                         Cantidad = Cantidad.Replace(",", ".");
                                         break;
                                     case (","):
                                         valor = valor.Replace(".", ",");
                                         Cantidad = Cantidad.Replace(".", ",");
                                         break;
                                 }
                             }

                             double valorRet = 0;
                             valorRet = (Convert.ToDouble(valor) * Convert.ToDouble(Cantidad)) * (prctRet / 100);

                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Cod_Ret").Cells.Item(pVal.Row).Specific;
                             oEdit.Value = CodRet;

                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Ret_Val").Cells.Item(pVal.Row).Specific;
                             oEdit.Value = valorRet.ToString();

                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Ret_Prct").Cells.Item(pVal.Row).Specific;
                             oEdit.Value = prctRet.ToString();
                         }
                         System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                         oRS = null;
                         GC.Collect();
                     }
                     System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                     item = null;
                     GC.Collect();
                 }
                 catch (Exception ex)
                 {
                     escribirLog("AddLine: " + ex.Message);
                     SBO_Application.MessageBox(ex.Message);
                 }
             }*/

            /* if (((pVal.ItemUID == "38") && (pVal.ColUID == "11") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) && (pVal.Before_Action == false)))

                 {
                 string codigoArti = "";
                 string WtLiable;
                 try
                 {
                     oForm = SBO_Application.Forms.Item(FormUID);

                     oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                     oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific;
                     codigoArti = oEdit.Value;

                     SAPbobsCOM.Items item = (SAPbobsCOM.Items)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                     item.GetByKey(codigoArti);
                     WtLiable = item.UserFields.Fields.Item("U_SCL_WTLiable").Value;

                     if (WtLiable == "Y")
                     {
                         string sSQL = "";
                         Recordset oRS;
                         oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                         sSQL = string.Format(Consultas.Default.CulculoRet, codigoArti);
                         escribirLog("ConsultaRet: " + sSQL);
                         oRS.DoQuery(sSQL);
                         if (oRS.RecordCount > 0)
                         {
                             string CodRet = "";
                             string Cantidad = "";
                             string valor = "";
                             double prctRet = 0;

                             CodRet = oRS.Fields.Item("WTCode").Value;

                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific;
                             Cantidad = oEdit.Value;

                             SAPbobsCOM.WithholdingTaxCodes oWithholding;
                             oWithholding = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
                             oWithholding.GetByKey(CodRet);
                             prctRet = oWithholding.BaseAmount;

                             oEdit = oMatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific;
                             valor = oEdit.Value;
                             valor = Regex.Replace(valor, "[$|COP|USD|EUR]", "");
                             valor = valor.Replace(DatosGlobalesFE.sapMillar, "");

                             if (String.IsNullOrEmpty(valor))
                             {
                                 valor = "0";
                             }

                             if (String.IsNullOrEmpty(Cantidad))
                             {
                                 Cantidad = "0";
                             }

                             if (!valor.Contains(DatosGlobalesFE.localdecimal) || !Cantidad.Contains(DatosGlobalesFE.localdecimal))
                             {
                                 switch (DatosGlobalesFE.localdecimal)
                                 {
                                     case ("."):
                                         valor = valor.Replace(",", ".");
                                         Cantidad = Cantidad.Replace(",", ".");
                                         break;
                                     case (","):
                                         valor = valor.Replace(".", ",");
                                         Cantidad = Cantidad.Replace(".", ",");
                                         break;
                                 }
                             }

                             double valorRet = 0;
                             valorRet = (Convert.ToDouble(valor) * Convert.ToDouble(Cantidad)) * (prctRet / 100);

                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Cod_Ret").Cells.Item(pVal.Row).Specific;
                             oEdit.Value = CodRet;

                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Ret_Val").Cells.Item(pVal.Row).Specific;
                             oEdit.Value = valorRet.ToString();

                             oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Ret_Prct").Cells.Item(pVal.Row).Specific;
                             oEdit.Value = prctRet.ToString();
                         }
                         System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                         oRS = null;
                         GC.Collect();
                     }
                     System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                     item = null;
                     GC.Collect();
                 }
                 catch (Exception ex)
                 {
                     escribirLog("ChangeQuantity: " + ex.Message);
                     SBO_Application.MessageBox(ex.Message);
                 }
             }*/

            if (((pVal.ItemUID == "38") && (pVal.ColUID == "U_SCL_Cod_Ret") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) && (pVal.Before_Action == false)))
            {
                string RetArt = "";
                string codigoArti = "";

                try
                {
                    string sSQL = "";
                    oForm = SBO_Application.Forms.Item(FormUID);
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                    oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific;
                    codigoArti = oEdit.Value;
                    oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Cod_Ret").Cells.Item(pVal.Row).Specific;
                    RetArt = oEdit.Value;

                    Recordset oRS;
                    oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //sSQL = string.Format(Consultas.Default.CntsNvl2, RetArt, codigoArti);
                    sSQL = string.Format(Consultas.Default.CalculoRet, codigoArti);
                    escribirLog("ConsultaRet: " + sSQL);
                    oRS.DoQuery(sSQL);
                    if (oRS.RecordCount > 0)
                    {
                        string CodRet = "";
                        string Cantidad = "";
                        string valor = "";
                        double prctRet = 0;

                        CodRet = oRS.Fields.Item("WTCode").Value;

                        oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific;
                        Cantidad = oEdit.Value;

                        SAPbobsCOM.WithholdingTaxCodes oWithholding;
                        oWithholding = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
                        oWithholding.GetByKey(CodRet);
                        prctRet = oWithholding.BaseAmount;


                        oEdit = oMatrix.Columns.Item("21").Cells.Item(pVal.Row).Specific;
                        valor = oEdit.Value;
                        valor = Regex.Replace(valor, "[$|COP|USD|EUR]", "");
                        valor = valor.Replace(DatosGlobalesFE.sapMillar, "");

                        if (String.IsNullOrEmpty(valor))
                        {
                            valor = "0";
                        }

                        if (String.IsNullOrEmpty(Cantidad))
                        {
                            Cantidad = "0";
                        }

                        if (!valor.Contains(DatosGlobalesFE.localdecimal) || !Cantidad.Contains(DatosGlobalesFE.localdecimal))
                        {
                            switch (DatosGlobalesFE.localdecimal)
                            {
                                case ("."):
                                    valor = valor.Replace(",", ".");
                                    Cantidad = Cantidad.Replace(",", ".");
                                    break;
                                case (","):
                                    valor = valor.Replace(".", ",");
                                    Cantidad = Cantidad.Replace(".", ",");
                                    break;
                            }
                        }

                        double valorRet = 0;
                        valorRet = (Convert.ToDouble(valor) * Convert.ToDouble(Cantidad)) * (prctRet / 100);

                        oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Ret_Val").Cells.Item(pVal.Row).Specific;
                        oEdit.Value = valorRet.ToString();

                        oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Ret_Prct").Cells.Item(pVal.Row).Specific;
                        oEdit.Value = prctRet.ToString();
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                    oRS = null;
                    GC.Collect();
                }
                catch (Exception ex)
                {
                    escribirLog("ChangePrice: " + ex.Message);
                    SBO_Application.MessageBox(ex.Message);
                }
            }

            if (((pVal.ItemUID == "38") && (pVal.ColUID == "14") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) && (pVal.Before_Action == false)))
            {
                string codigoArti = "";
                string WtLiable;
                try
                {
                    oForm = SBO_Application.Forms.Item(FormUID);

                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                    oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific;
                    codigoArti = oEdit.Value;

                    SAPbobsCOM.Items item = (SAPbobsCOM.Items)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                    item.GetByKey(codigoArti);
                    WtLiable = item.UserFields.Fields.Item("U_SCL_WTLiable").Value;

                    if (WtLiable == "Y")
                    {
                        string sSQL = "";
                        Recordset oRS;
                        oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        sSQL = string.Format(Consultas.Default.CulculoRet, codigoArti);
                        escribirLog("ConsultaRet: " + sSQL);
                        oRS.DoQuery(sSQL);
                        if (oRS.RecordCount > 0)
                        {
                            string CodRet = "";
                            string Cantidad = "";
                            string valor = "";
                            double prctRet = 0;

                            CodRet = oRS.Fields.Item("WTCode").Value;

                            oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific;
                            Cantidad = oEdit.Value;

                            SAPbobsCOM.WithholdingTaxCodes oWithholding;
                            oWithholding = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
                            oWithholding.GetByKey(CodRet);
                            prctRet = oWithholding.BaseAmount;


                            oEdit = oMatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific;
                            valor = oEdit.Value;
                            valor = Regex.Replace(valor, "[$|COP|USD|EUR]", "");
                            valor = valor.Replace(DatosGlobalesFE.sapMillar, "");

                            if (String.IsNullOrEmpty(valor))
                            {
                                valor = "0";
                            }

                            if (String.IsNullOrEmpty(Cantidad))
                            {
                                Cantidad = "0";
                            }

                            if (!valor.Contains(DatosGlobalesFE.localdecimal) || !Cantidad.Contains(DatosGlobalesFE.localdecimal))
                            {
                                switch (DatosGlobalesFE.localdecimal)
                                {
                                    case ("."):
                                        valor = valor.Replace(",", ".");
                                        Cantidad = Cantidad.Replace(",", ".");
                                        break;
                                    case (","):
                                        valor = valor.Replace(".", ",");
                                        Cantidad = Cantidad.Replace(".", ",");
                                        break;
                                }
                            }

                            double valorRet = 0;
                            valorRet = (Convert.ToDouble(valor) * Convert.ToDouble(Cantidad)) * (prctRet / 100);

                            oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Cod_Ret").Cells.Item(pVal.Row).Specific;
                            oEdit.Value = CodRet;

                            oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Ret_Val").Cells.Item(pVal.Row).Specific;
                            oEdit.Value = valorRet.ToString();

                            oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Ret_Prct").Cells.Item(pVal.Row).Specific;
                            oEdit.Value = prctRet.ToString();
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                        oRS = null;
                        GC.Collect();
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                    item = null;
                    GC.Collect();
                }
                catch (Exception ex)
                {
                    escribirLog("ChangePrice: " + ex.Message);
                    SBO_Application.MessageBox(ex.Message);
                }
            }

            if (pVal.ItemUID == "1" && pVal.FormMode == 3 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == true)
            {
                try
                {
                    string CardCode = "";
                    oForm = SBO_Application.Forms.Item(FormUID);

                    oEdit = oForm.Items.Item("4").Specific;
                    CardCode = oEdit.Value;

                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;

                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Columns.Add(new System.Data.DataColumn("1", typeof(string)));
                    dt.Columns.Add(new System.Data.DataColumn("2", typeof(double)));
                    dt.Columns.Add(new System.Data.DataColumn("3", typeof(double)));

                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {
                        string codRet = "";
                        string Valor = "";
                        string price = "";
                        string cantidad = "";

                        oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Cod_Ret").Cells.Item(i).Specific;
                        codRet = oEdit.Value;

                        if (!string.IsNullOrEmpty(codRet))
                        {
                            DataRow newrow = dt.NewRow();
                            newrow[0] = codRet;

                            oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SCL_Ret_Val").Cells.Item(i).Specific;
                            Valor = oEdit.Value;

                            oEdit = oMatrix.Columns.Item("14").Cells.Item(i).Specific;
                            price = oEdit.Value;
                            price = Regex.Replace(price, "[$|COP|USD|EUR]", "");
                            price = price.Replace(DatosGlobalesFE.sapMillar, "");

                            oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(i).Specific;
                            cantidad = oEdit.Value;

                            if (String.IsNullOrEmpty(price))
                            {
                                price = "0";
                            }

                            if (String.IsNullOrEmpty(cantidad))
                            {
                                cantidad = "0";
                            }

                            if (!Valor.Contains(DatosGlobalesFE.localdecimal) || !price.Contains(DatosGlobalesFE.localdecimal) || !cantidad.Contains(DatosGlobalesFE.localdecimal))
                            {
                                switch (DatosGlobalesFE.localdecimal)
                                {
                                    case ("."):
                                        Valor = Valor.Replace(",", ".");
                                        price = price.Replace(",", ".");
                                        cantidad = cantidad.Replace(",", ".");
                                        break;
                                    case (","):
                                        Valor = Valor.Replace(".", ",");
                                        price = price.Replace(".", ",");
                                        cantidad = cantidad.Replace(".", ",");
                                        break;
                                }
                            }
                            double valorBase = 0;
                            valorBase = Convert.ToDouble(price) * Convert.ToDouble(cantidad);
                            escribirLog("U_SCL_Ret_Val: " + Valor);
                            escribirLog("valorbase: " + valorBase);
                            newrow[1] = Valor;
                            newrow[2] = valorBase;
                            dt.Rows.Add(newrow);
                        }
                    }
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        oForm = SBO_Application.Forms.Item(FormUID);

                        SBO_Application.Menus.Item("5897").Activate();

                        var query = from r in dt.AsEnumerable()
                                    group r by r.Field<string>(0) into groupedTable
                                    select new
                                    {
                                        id = groupedTable.Key,
                                        sumOfValue = groupedTable.Sum(s => s.Field<double>("2")),
                                        valueBase = groupedTable.Sum(s => s.Field<double>("3"))
                                    };

                        System.Data.DataTable newDt = ConvertToDataTable(query);

                        oForm = SBO_Application.Forms.ActiveForm;
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("6").Specific;

                        oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("68").Cells.Item(1).Specific;
                        string docBase = oEdit.Value;

                        if (docBase == "-1")
                        {
                            int rowmatrix = 1;
                            for (int i = 0; i < newDt.Rows.Count;)
                            {
                                bool existRet = false;
                                SAPbobsCOM.WithholdingTaxCodes oWithholding;
                                oWithholding = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
                                existRet = oWithholding.GetByKey(newDt.Rows[i]["id"].ToString());

                                if (!string.IsNullOrEmpty(newDt.Rows[i]["id"].ToString()) && existRet)
                                {
                                    oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(rowmatrix).Specific;
                                    escribirLog("codRet: " + newDt.Rows[i]["id"].ToString());
                                    oEdit.Value = newDt.Rows[i]["id"].ToString();

                                    oColunm = oMatrix.Columns.Item("14");
                                    if (oColunm.Editable == true)
                                    {
                                        oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("14").Cells.Item(rowmatrix).Specific;
                                        escribirLog("SumValorRet: " + newDt.Rows[i]["sumOfValue"].ToString());
                                        oEdit.Value = newDt.Rows[i]["sumOfValue"].ToString();
                                    }
                                    else
                                    {
                                        oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("28").Cells.Item(rowmatrix).Specific;
                                        escribirLog("SumValorRet: " + newDt.Rows[i]["sumOfValue"].ToString());
                                        oEdit.Value = newDt.Rows[i]["sumOfValue"].ToString();
                                    }

                                    oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(rowmatrix).Specific;
                                    if (string.IsNullOrEmpty(oEdit.Value))
                                    {
                                        string msg = "";
                                        msg = "Retencion " + newDt.Rows[i]["id"].ToString() + " no asignada al socio de negocio";
                                        SBO_Application.StatusBar.SetSystemMessage(msg, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
                                    }
                                    else
                                    {
                                        rowmatrix++;
                                    }
                                }
                                if (!string.IsNullOrEmpty(newDt.Rows[i]["id"].ToString()) && existRet == false)
                                {
                                    string msg = "";
                                    msg = "La Retencion " + newDt.Rows[i]["id"].ToString() + " no existe";
                                    SBO_Application.StatusBar.SetSystemMessage(msg, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
                                }

                                i++;

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oWithholding);
                                oWithholding = null;
                                GC.Collect();
                            }
                        }
                        else
                        {
                            for (int i = 1; i <= oMatrix.RowCount; i++)
                            {
                                bool existRet = false;
                                string retActual = "";
                                oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific;
                                retActual = oEdit.Value;
                                for (int j = 0; j < newDt.Rows.Count; j++)
                                {
                                    string retNueva = "";
                                    retNueva = newDt.Rows[j]["id"].ToString();
                                    if (retActual == retNueva)
                                    {
                                        existRet = true;
                                        oColunm = oMatrix.Columns.Item("14");
                                        if (oColunm.Editable == true)
                                        {
                                            oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("14").Cells.Item(i).Specific;
                                            escribirLog("SumValorRet: " + newDt.Rows[j]["sumOfValue"].ToString());
                                            oEdit.Value = newDt.Rows[j]["sumOfValue"].ToString();
                                        }
                                        else
                                        {
                                            oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("28").Cells.Item(i).Specific;
                                            escribirLog("SumValorRet: " + newDt.Rows[j]["sumOfValue"].ToString());
                                            oEdit.Value = newDt.Rows[j]["sumOfValue"].ToString();
                                        }
                                    }
                                }
                                if (existRet == false)
                                {
                                    oMatrix.DeleteRow(i);
                                }
                            }
                        }

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            oItem = oForm.Items.Item("1");
                            oItem.Click();
                        }
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            oItem = oForm.Items.Item("1");
                            oItem.Click();
                        }
                    }
                }
                catch (Exception ex)
                {
                    escribirLog("CalcuRet: " + ex.Message);
                    SBO_Application.MessageBox(ex.Message);
                }
            }

            if ((pVal.FormType == 141) && ((pVal.ItemUID == "38") && (pVal.ColUID == "160") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) && (pVal.Before_Action == false)))
            {
                try
                {
                    string TaxCode = "";

                    oForm = SBO_Application.Forms.Item(FormUID);
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                    oEdit = oMatrix.Columns.Item("160").Cells.Item(pVal.Row).Specific;
                    TaxCode = oEdit.Value;
                    SAPbobsCOM.UserTables tablas = null;
                    SAPbobsCOM.UserTable tabla = null;

                    tablas = oCompany.UserTables;
                    tabla = tablas.Item("SCL_IVA_MAYOR");
                    if (tabla.GetByKey(TaxCode))
                    {
                        string valor = "";
                        SAPbobsCOM.SalesTaxCodes oTaxCode;
                        oTaxCode = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesTaxCodes);
                        oTaxCode.GetByKey(TaxCode);

                        oEdit = oMatrix.Columns.Item("21").Cells.Item(pVal.Row).Specific;
                        valor = oEdit.Value;
                        valor = Regex.Replace(valor, "[$|COP|USD|EUR]", "");
                        valor = valor.Replace(DatosGlobalesFE.sapMillar, "");

                        if (!valor.Contains(DatosGlobalesFE.localdecimal))
                        {
                            switch (DatosGlobalesFE.localdecimal)
                            {
                                case ("."):
                                    valor = valor.Replace(",", ".");
                                    break;
                                case (","):
                                    valor = valor.Replace(".", ",");
                                    break;
                            }
                        }

                        double valorgasto = 0;
                        valorgasto = Convert.ToDouble(valor) * (oTaxCode.Rate / 100);
                        oComboBox = oMatrix.Columns.Item("111").Cells.Item(pVal.Row).Specific;
                        oComboBox.Select(tabla.UserFields.Fields.Item("U_SCL_GastoAdi").Value);

                        oEdit = oMatrix.Columns.Item("112").Cells.Item(pVal.Row).Specific;
                        oEdit.Value = valorgasto.ToString();

                        oComboBox = oMatrix.Columns.Item("115").Cells.Item(pVal.Row).Specific;
                        oComboBox.Select(tabla.UserFields.Fields.Item("U_SCL_GastoAdi").Value);

                        oEdit = oMatrix.Columns.Item("116").Cells.Item(pVal.Row).Specific;
                        oEdit.Value = "-" + valorgasto.ToString();

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oTaxCode);
                        oTaxCode = null;
                        GC.Collect();
                    }
                }
                catch (Exception ex)
                {
                    escribirLog("IndicadorImpu: " + ex.Message);
                    SBO_Application.MessageBox(ex.Message);
                }
            }

            if ((pVal.FormType == 181) && ((pVal.ItemUID == "38") && (pVal.ColUID == "160") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) && (pVal.Before_Action == false)))
            {
                try
                {
                    string TaxCode = "";

                    oForm = SBO_Application.Forms.Item(FormUID);
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                    oEdit = oMatrix.Columns.Item("160").Cells.Item(pVal.Row).Specific;
                    TaxCode = oEdit.Value;
                    SAPbobsCOM.UserTables tablas = null;
                    SAPbobsCOM.UserTable tabla = null;

                    tablas = oCompany.UserTables;
                    tabla = tablas.Item("SCL_IVA_MAYOR");
                    if (tabla.GetByKey(TaxCode))
                    {
                        string valor = "";
                        SAPbobsCOM.SalesTaxCodes oTaxCode;
                        oTaxCode = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesTaxCodes);
                        oTaxCode.GetByKey(TaxCode);

                        oEdit = oMatrix.Columns.Item("21").Cells.Item(pVal.Row).Specific;
                        valor = oEdit.Value;
                        valor = Regex.Replace(valor, "[$|COP|USD|EUR]", "");
                        valor = valor.Replace(DatosGlobalesFE.sapMillar, "");

                        if (!valor.Contains(DatosGlobalesFE.localdecimal))
                        {
                            switch (DatosGlobalesFE.localdecimal)
                            {
                                case ("."):
                                    valor = valor.Replace(",", ".");
                                    break;
                                case (","):
                                    valor = valor.Replace(".", ",");
                                    break;
                            }
                        }

                        double valorgasto = 0;
                        valorgasto = Convert.ToDouble(valor) * (oTaxCode.Rate / 100);
                        oComboBox = oMatrix.Columns.Item("111").Cells.Item(pVal.Row).Specific;
                        oComboBox.Select(tabla.UserFields.Fields.Item("U_SCL_GastoAdi").Value);

                        oEdit = oMatrix.Columns.Item("112").Cells.Item(pVal.Row).Specific;
                        oEdit.Value = valorgasto.ToString();

                        oComboBox = oMatrix.Columns.Item("115").Cells.Item(pVal.Row).Specific;
                        oComboBox.Select(tabla.UserFields.Fields.Item("U_SCL_GastoAdi").Value);

                        oEdit = oMatrix.Columns.Item("116").Cells.Item(pVal.Row).Specific;
                        oEdit.Value = "-" + valorgasto.ToString();

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oTaxCode);
                        oTaxCode = null;
                        GC.Collect();
                    }
                }
                catch (Exception ex)
                {
                    escribirLog("IndicadorImpu: " + ex.Message);
                    SBO_Application.MessageBox(ex.Message);
                }
            }
        }

        /// <summary>
        /// Sumar datos de data table agrupados por codigo de retencion
        /// </summary>
        public System.Data.DataTable ConvertToDataTable<T>(IEnumerable<T> varlist)
        {
            System.Data.DataTable dtReturn = new System.Data.DataTable();


            // column names
            PropertyInfo[] oProps = null;


            if (varlist == null) return dtReturn;


            foreach (T rec in varlist)
            {
                // Use reflection to get property names, to create table, Only first time, others will follow
                if (oProps == null)
                {
                    oProps = ((Type)rec.GetType()).GetProperties();
                    foreach (PropertyInfo pi in oProps)
                    {
                        Type colType = pi.PropertyType;


                        if ((colType.IsGenericType) && (colType.GetGenericTypeDefinition() == typeof(Nullable<>)))
                        {
                            colType = colType.GetGenericArguments()[0];
                        }


                        dtReturn.Columns.Add(new System.Data.DataColumn(pi.Name, colType));
                    }
                }


                DataRow dr = dtReturn.NewRow();


                foreach (PropertyInfo pi in oProps)
                {
                    dr[pi.Name] = pi.GetValue(rec, null) == null ? DBNull.Value : pi.GetValue
                    (rec, null);
                }


                dtReturn.Rows.Add(dr);
            }
            return dtReturn;
        }

        /// <summary>
        /// Manaejo de eventos folder finanzas en maestro de articulos
        /// </summary>
        /// <param name="pVal">Tipo de evento</param>
        /// <param name="formUID">Identificador del formulario</param>
        private void EventFieldFinanza(ItemEvent pVal, string FormUID)
        {
            if (pVal.ItemUID == "FINANZA001" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.Before_Action == true)
            {
                try
                {
                    oForm = SBO_Application.Forms.Item(FormUID);
                    oForm.Select();
                    oForm.PaneLevel = 99;
                    oForm.Mode = BoFormMode.fm_OK_MODE;
                    //oItem = oForm.Items.Item("Usuario");
                    //oItem.Click();
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("LoadItem: " + ex.Message);
                }
            }

            if (pVal.ItemUID == "RET01" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.Before_Action == true)
            {
                try
                {
                    bool wtactivo;
                    oForm = SBO_Application.Forms.Item(FormUID);
                    oItem = oForm.Items.Item("RET01");
                    oChekBox = (CheckBox)oItem.Specific;
                    wtactivo = oChekBox.Checked;
                    if (wtactivo)
                    {
                        oItem = oForm.Items.Item("RET02");
                        oItem.Visible = true;

                        oItem = oForm.Items.Item("RET03");
                        oItem.Visible = true;
                    }
                    else
                    {
                        oItem = oForm.Items.Item("RET02");
                        oItem.Visible = false;

                        oItem = oForm.Items.Item("RET03");
                        oItem.Visible = false;
                    }
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("LoadItem: " + ex.Message);
                }
            }

            if (pVal.ItemUID == "RET03" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.Before_Action == true)
            {
                try
                {
                    oForm = SBO_Application.Forms.Item(FormUID);
                    oItem = oForm.Items.Item("5");
                    oEdit = (EditText)oItem.Specific;

                    CodigoArticulo = oEdit.Value;
                    CrearFormularioRet(pVal, FormUID, CodigoArticulo);
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("LoadItem: " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Manaejo de eventos items socio de negocios
        /// </summary>
        /// <param name="pVal">Tipo de evento</param>
        /// <param name="formUID">Identificador del formulario</param>
        private void EventFieldSocio(ItemEvent pVal, string FormUID)
        {
            //digito verificacion directamente del campo

            if (pVal.ItemUID == "41" && pVal.EventType == BoEventTypes.et_LOST_FOCUS && pVal.Before_Action == false)
            {
                try
                {
                    string CardCode = null;
                    oForm = SBO_Application.Forms.ActiveForm;

                    oEdit = ((SAPbouiCOM.EditText)(oForm.Items.Item("41").Specific));
                    CardCode = oEdit.Value;
                    int tam_var = CardCode.Length;
                    string Var_Sub = CardCode.Substring((tam_var - 1), 1);
                    string nitt = CardCode.Split('-')[1];
                    string nit = CardCode.ToUpper();
                    if (int.Parse(nitt.ToString()) > 0)
                    {
                        string verifiNIT = verinit(CardCode);
                        if (verifiNIT == nitt)
                        {

                        }
                        else
                        {
                            throw new Exception("El codigo de verificacion es erroneo");

                        }
                    }

                }
                catch (Exception ex)
                {
                    SBO_Application.SetStatusBarMessage("El codigo de verificacion es erroneo");
                    //SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("LoadRetSN: " + ex.Message);

                }
            }
            if (pVal.ItemUID == "RETSN02" && pVal.EventType == BoEventTypes.et_CLICK && pVal.Before_Action == true)
            {
                try
                {
                    oForm = SBO_Application.Forms.Item(FormUID);
                    oItem = oForm.Items.Item("5");
                    oEdit = (EditText)oItem.Specific;

                    CodigoSocio = oEdit.Value;
                    CrearFormularioAutoRet(pVal, FormUID, CodigoSocio);
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("LoadRetSN: " + ex.Message);
                }
            }

            //verificar campos exogena

            if ((pVal.ItemUID == "EXO03" || pVal.ItemUID == "EXO05" || pVal.ItemUID == "EXO07" || pVal.ItemUID == "EXO09") && pVal.EventType == BoEventTypes.et_LOST_FOCUS && pVal.Before_Action == false)
            {
                //  string error = "";
                try
                {
                    //string NIT = null;
                    string apellido1 = null;
                    string apellido2 = null;
                    string nombre = null;
                    string nomadic = null;
                    //string razonSocial = null;
                    //string codMun = null;
                    //string codPais= null;
                    oForm = SBO_Application.Forms.ActiveForm;


                    //oEdit = ((SAPbouiCOM.EditText)(oForm.Items.Item("EXO01").Specific));
                    //NIT = oEdit.Value;
                    oEdit = ((SAPbouiCOM.EditText)(oForm.Items.Item("EXO03").Specific));
                    apellido1 = oEdit.Value;
                    oEdit = ((SAPbouiCOM.EditText)(oForm.Items.Item("EXO05").Specific));
                    apellido2 = oEdit.Value;
                    oEdit = ((SAPbouiCOM.EditText)(oForm.Items.Item("EXO07").Specific));
                    nombre = oEdit.Value;
                    oEdit = ((SAPbouiCOM.EditText)(oForm.Items.Item("EXO09").Specific));
                    nomadic = oEdit.Value;
                    //oEdit = ((SAPbouiCOM.EditText)(oForm.Items.Item("EXO17").Specific));
                    ////razonSocial = oEdit.Value;
                    //oEdit = ((SAPbouiCOM.EditText)(oForm.Items.Item("EXO19").Specific));
                    //codMun = oEdit.Value;
                    //oEdit = ((SAPbouiCOM.EditText)(oForm.Items.Item("EXO22").Specific));
                    //codPais = oEdit.Value;
                    /*oComboBox = ((SAPbouiCOM.ComboBox)(oForm.Items.Item("EXO09").Specific));
                    TipoDoc = oEdit.Value;
                    
                    oItem = oForm.Items.Item("EXO11");
                    //oComboBox = (ComboBox)oItem.Specific;
                    TipoDoc = oComboBox.Item.Specific;

                    oComboBox = (SAPbouiCOM.ComboBox)oNewItem.Specific;
                    oComboBox.DataBind.SetBound(true, "OCRD", "U_SCL_TipoDoc");

                    TipoDoc = oEdit.Value;
                    oEdit = ((SAPbouiCOM.EditText)(oForm.Items.Item("EXO013").Specific));
                    TipPers = oEdit.Value;*/
                    //if (string.IsNullOrEmpty(NIT))
                    //{
                    //    throw new Exception("El campo NIT o C.C. se encuentra vacio");
                    //}
                    if (string.IsNullOrEmpty(apellido1))
                    {
                        throw new Exception("El campo Primer Apellido se encuentra vacio");
                    }
                    else if (string.IsNullOrEmpty(apellido2))
                    {
                        throw new Exception("El campo Segundo Apellido se encuentra vacio");
                    }
                    else if (string.IsNullOrEmpty(nombre))
                    {
                        throw new Exception("El campo Primer Nombre se encuentra vacio");
                    }
                    else if (string.IsNullOrEmpty(nomadic))
                    {
                        throw new Exception("El campo Nombres Adicionales se encuentra vacio");
                    }
                    //else if (string.IsNullOrEmpty(razonSocial))
                    //{
                    //    throw new Exception("El campos Razon Social se encuentra vacio");
                    //}
                    //else if (string.IsNullOrEmpty(codMun))
                    //{
                    //    throw new Exception("El campo Municipio MM se encuentra vacio");
                    //}
                    //else if (string.IsNullOrEmpty(codPais))
                    //{
                    //    throw new Exception("El campo Pais Domicilió se encuentra vacio");
                    //}

                }
                catch (Exception ex)
                {
                    SBO_Application.SetStatusBarMessage(ex.Message);
                    //SBO_Application.MessageBox(ex.Message);
                     Business.escribirLog("LoadRetSN: " + ex.Message);

                }
            }
            // campos exogena 
            if (pVal.ItemUID == "EXOG001" && pVal.EventType == BoEventTypes.et_CLICK && pVal.Before_Action == true)
            {
                try
                {
                    oForm = SBO_Application.Forms.Item(FormUID);
                    oForm.Select();
                    oForm.PaneLevel = 99;
                    //oForm.Mode = BoFormMode.fm_OK_MODE;
                    oItem = oForm.Items.Item("EXG04");

                    oItem.Click();
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("CampExog: " + ex.Message);
                }
            }
        }

        private void EventFieldDS(ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.Item oItem = null;
            // oForm = SBO_Application.Forms.Item(FormUID);
            if (pVal.ItemUID == "CRED_SL" && pVal.EventType == BoEventTypes.et_CLICK && pVal.Before_Action == true)
            {
                try
                {
                    oForm = SBO_Application.Forms.Item(FormUID);
                    oForm.Select();
                    oForm.PaneLevel = 99;
                    oForm.Mode = BoFormMode.fm_OK_MODE;
                    oItem = oForm.Items.Item("CRD06");

                    oItem.Click();
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("CamposDS: " + ex.Message);
                }
            }

            if ((pVal.FormMode == 2 && (pVal.ItemUID == "1") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK) && (pVal.Before_Action == false)))
            {

                try
                {
                    oForm = SBO_Application.Forms.Item(FormUID);
                    oItem = oForm.Items.Item("CRD07");
                    oEdit = oItem.Specific;

                    if (string.IsNullOrEmpty(oEdit.Value))
                    {

                        oForm = SBO_Application.Forms.Item(FormUID);
                        oItem = oForm.Items.Item("CRD05");
                        oEdit = oItem.Specific;
                        DatosGlobServiceLayer.password = oEdit.Value;
                        string claveEncriptado = oConnection.Encrypt(oEdit.Value);
                        oEdit.Value = claveEncriptado;

                        oForm = SBO_Application.Forms.Item(FormUID);
                        oItem = oForm.Items.Item("CRD07");
                        oEdit = oItem.Specific;
                        oEdit.Value = "Si";
                    }
                    GC.Collect();
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("CamposDS: " + ex.Message);
                }
            }
        }

        #region Eventos Transferencia de Stock Eurocares
        // Eventos Gestion de Series y Lotes
        private void EventFieldSeries(ItemEvent pVal, string FormUID)
        {
            try
            {
                if (pVal.FormTypeEx == "25" && pVal.ItemUID == "22" && pVal.EventType == BoEventTypes.et_KEY_DOWN && pVal.CharPressed == 9 && pVal.Before_Action == false)
                {
                    oForm = SBO_Application.Forms.Item(FormUID);
                    oItem = oForm.Items.Item("8");
                    oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {

                        oItem = oForm.Items.Item("1");
                        oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                        oItem = oForm.Items.Item("3");
                        oMatrix = oItem.Specific;
                        //oMatrix.IsRowSelected Se puede manear como un ciclo

                        //OCellPosition = oMatrix.GetCellFocus();
                        //int position = OCellPosition.rowIndex;
                        oEdit = oMatrix.Columns.Item("4").Cells.Item(contSeries).Specific;//cANTIDAD
                        string cantidad = oEdit.Value.ToString();

                        oEdit = oMatrix.Columns.Item("5").Cells.Item(contSeries).Specific;//TOTAL SELECCIONADO
                        string seleccionado = oEdit.Value.ToString();

                        if (seleccionado == cantidad)
                        {
                            contSeries++;
                            oMatrix.Columns.Item("4").Cells.Item(contSeries).Click();

                        }
                    }
                }
                //GC.Collect();
                if (pVal.FormTypeEx == "25" && pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                {

                    oForm = SBO_Application.Forms.Item(FormUID);
                    oItem = oForm.Items.Item("3");
                    oMatrix = oItem.Specific;
                    contSeries = 1;
                    while (!oMatrix.IsRowSelected(contSeries))
                    {
                        contSeries++;
                    }
                }
            }
            catch (Exception ex)
            {
                //SBO_Application.MessageBox(ex.Message);
                escribirLog("EventSeries: " + ex.Message);
            }
        }

        // Eventos Facturar Traslados        
        private void CrearItemsTS(ItemEvent pVal, string FormUID)
        {
            if (pVal.Before_Action == true)
            {
                try
                {
                    oForm = SBO_Application.Forms.Item(FormUID);
                    oForm.DataSources.UserDataSources.Add("dtschkDev", BoDataType.dt_SHORT_TEXT, 20);

                    oItem = oForm.Items.Item("1250000074");

                    // get the event sending form
                    oNewItem = oForm.Items.Add("BtnFactura", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    oNewItem.ToPane = 0;
                    oNewItem.FromPane = 0;
                    //oNewItem.Left = 541;
                    //oNewItem.Left = 320;
                    ////oNewItem.Top = 470;
                    //oNewItem.Top = 438;
                    //oNewItem.Height = 19;
                    //oNewItem.Width = 122;
                    oNewItem.Left = oItem.Left - 120;
                    oNewItem.Top = oItem.Top;
                    oNewItem.Height = oItem.Height;
                    oNewItem.Width = oItem.Width;
                    oNewItem.Visible = true;
                    oNewItem.AffectsFormMode = true;
                    oButton = ((oNewItem.Specific));
                    oButton.Caption = "Generar Factura";

                    oItem = oForm.Items.Item("10000053");

                    oNewItem = oForm.Items.Add("ChkDev", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    oNewItem.ToPane = 0;
                    oNewItem.FromPane = 0;
                    //oNewItem.Left = 168;
                    //oNewItem.Top = 181;
                    //oNewItem.Height = 14;
                    //oNewItem.Width = 122;
                    oNewItem.Left = oItem.Left + 160;
                    oNewItem.Top = oItem.Top + 25;
                    oNewItem.Height = oItem.Height;
                    oNewItem.Width = oItem.Width;
                    oNewItem.Visible = true;
                    oNewItem.AffectsFormMode = true;
                    oChekBox = ((oNewItem.Specific));
                    oChekBox.DataBind.SetBound(true, "OWTR", "U_SCL_ChkDev");
                    //oChekBox.DataBind.SetBound(true, "", "dtschkDev");

                    oNewItem = oForm.Items.Add("StcDev", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    oNewItem.ToPane = 0;
                    oNewItem.FromPane = 0;
                    //oNewItem.Left = 0;
                    //oNewItem.Top = 181;
                    //oNewItem.Height = 14;
                    //oNewItem.Width = 169;
                    oNewItem.Left = oItem.Left;
                    oNewItem.Top = oItem.Top + 25;
                    oNewItem.Height = oItem.Height;
                    oNewItem.Width = oItem.Width;
                    oNewItem.Visible = true;
                    oNewItem.AffectsFormMode = true;
                    oStatic = ((StaticText)(oNewItem.Specific));
                    oStatic.Caption = "Devolución";

                }
                catch (Exception ex)
                {
                    escribirLog("CrearItemsTS:" + ex.Message);
                }
            }
        }
        // Eventos Documento Transferencia de Stock SAP

        //Comparar trasnferencias para facturar el resultado en una factura prelimnar ----
        private void EventFieldTranStock(ItemEvent pVal, string FormUID)
        {
            try
            {
                #region Boton Facturar
                if (pVal.FormTypeEx == "940" && pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "BtnFactura" && pVal.FormMode == 1 && pVal.Before_Action == false)
                {
                    string docEntryDev, docEntry, SNFactura, sSQL, docPrel = null;
                    int numDoc = 0;
                    bool dev = false;

                    List<lineas_Factura> traslado = new List<lineas_Factura>();                    List<lineas_Factura> trasladoDev = new List<lineas_Factura>();
                    Factura_Traslado cab = new Factura_Traslado();                    List<lineas_Factura> FacturarLin = new List<lineas_Factura>();

                    //SAPbobsCOM.StockTransfer oStockTransfer = oCompany.GetBusinessObject(BoObjectTypes.oStockTransfer);                    //
                    Recordset oRecordSet = null;
                    Recordset oRecordSetLin = null;
                    oRecordSet = ((Recordset)(oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)));
                    oRecordSetLin = ((Recordset)(oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)));

                    oForm = SBO_Application.Forms.Item(FormUID);
                    oEdit = ((SAPbouiCOM.EditText)(oForm.Items.Item("11").Specific));
                    numDoc = Int32.Parse(oEdit.Value);

                    oItem = oForm.Items.Item("ChkDev");
                    oChekBox = oItem.Specific;

                    //Consultar DocEntry de la transsferencia relacionada como devolución
                    sSQL = string.Format(Consultas.Default.TransferenciaDev, numDoc);
                    oRecordSet.DoQuery(sSQL);
                    docEntry = Convert.ToString(oRecordSet.Fields.Item("DocOrg").Value);
                    docEntryDev = Convert.ToString(oRecordSet.Fields.Item("DocDev").Value);
                    SNFactura = Convert.ToString(oRecordSet.Fields.Item("SN").Value);
                    docPrel = Convert.ToString(oRecordSet.Fields.Item("DocPreliminar").Value);

                    if (docPrel != "0")
                    {
                        SBO_Application.StatusBar.SetText("Existe un documento preliminar de la transferencia número " + numDoc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        Business.escribirLog("ValidarDocPreliminar: Existe un documento preliminar de la transferencia número " + numDoc);
                        return;
                    }
                    if (!string.IsNullOrEmpty(SNFactura))
                    {
                        if (oChekBox.Checked == true)
                        {
                            try
                            {
                                if (!docEntry.Equals("0"))
                                {
                                    lineas_Factura lin = new lineas_Factura();
                                    Dictionary<string, string> s = new Dictionary<string, string>();
                                    List<string> l = new List<string>();

                                    sSQL = string.Format(Consultas.Default.TransferenciaCab, docEntry);
                                    oRecordSet.DoQuery(sSQL);

                                    cab = new Factura_Traslado();
                                    cab.SN = SNFactura;
                                    cab.fechaPed = Convert.ToString(oRecordSet.Fields.Item("Fecha Ped.").Value);
                                    cab.fechaCir = Convert.ToString(oRecordSet.Fields.Item("Fecha Cir.").Value);
                                    cab.nombreMed = Convert.ToString(oRecordSet.Fields.Item("Nombre Med.").Value);
                                    cab.nombrePac = Convert.ToString(oRecordSet.Fields.Item("Nombre Pac.").Value);
                                    cab.clinica = Convert.ToString(oRecordSet.Fields.Item("Clinica").Value);
                                    cab.cedula = Convert.ToString(oRecordSet.Fields.Item("Cedula").Value);
                                    cab.aFactura = Convert.ToString(oRecordSet.Fields.Item("Se factura.").Value);
                                    cab.ubicacion = Convert.ToString(oRecordSet.Fields.Item("AbsEntry").Value);
                                    cab.tipoDoc = Convert.ToString(oRecordSet.Fields.Item("Documento").Value);
                                    cab.transRef = docEntryDev;
                                    cab.transOrg = docEntry;
                                    // Datos Lineas
                                    sSQL = string.Format(Consultas.Default.Articulos_Series, docEntry);
                                    oRecordSetLin.DoQuery(sSQL);
                                    int numLinea = Convert.ToInt32(oRecordSetLin.Fields.Item("LineNum").Value);

                                    s = new Dictionary<string, string>();
                                    l = new List<string>();
                                    oRecordSetLin.MoveFirst();
                                    while (!oRecordSetLin.EoF)
                                    {
                                        if (Convert.ToInt32(oRecordSetLin.Fields.Item("LineNum").Value) == numLinea)
                                        {
                                            s.Add(Convert.ToString(oRecordSetLin.Fields.Item("Series").Value), Convert.ToString(oRecordSetLin.Fields.Item("Series").Value));
                                            l.Add(Convert.ToString(oRecordSetLin.Fields.Item("Lotes").Value));

                                            lin = new lineas_Factura()
                                            {
                                                codigo = Convert.ToString(oRecordSetLin.Fields.Item("ItemCode").Value),
                                                cantidad = Convert.ToInt32(oRecordSetLin.Fields.Item("Quantity").Value),
                                                almacen = Convert.ToString(oRecordSetLin.Fields.Item("WhsCode").Value),
                                                ciudad = Convert.ToString(oRecordSetLin.Fields.Item("OcrCode").Value),
                                                centroCostos = Convert.ToString(oRecordSetLin.Fields.Item("OcrCode2").Value),
                                                serie = s,
                                                lote = l
                                            };
                                            oRecordSetLin.MoveNext();
                                        }
                                        else if (!string.IsNullOrEmpty(lin.codigo))
                                        {
                                            traslado.Add(lin);
                                            numLinea++;
                                            lin = new lineas_Factura();
                                            s = new Dictionary<string, string>();
                                            l = new List<string>();
                                        }
                                        else
                                        {
                                            numLinea++;
                                        }
                                        if (oRecordSetLin.EoF)
                                        {
                                            traslado.Add(lin);
                                            s = new Dictionary<string, string>();
                                            l = new List<string>();
                                        }
                                        
                                    }

                                    // escribirLog("FacturarTraslado: Almaceno artículos del traslado de Stock origen");
                                    //Consultar articulos y series de la trasnferencia relacionada como devolución
                                    sSQL = string.Format(Consultas.Default.Articulos_Series, docEntryDev);
                                    oRecordSetLin.DoQuery(sSQL);
                                    oRecordSetLin.MoveFirst();
                                    numLinea = Convert.ToInt32(oRecordSetLin.Fields.Item("LineNum").Value);

                                    while (!oRecordSetLin.EoF)
                                    {
                                        if (Convert.ToInt32(oRecordSetLin.Fields.Item("LineNum").Value) == numLinea)
                                        {
                                            s.Add(Convert.ToString(oRecordSetLin.Fields.Item("Series").Value), Convert.ToString(oRecordSetLin.Fields.Item("Series").Value));
                                            l.Add(Convert.ToString(oRecordSetLin.Fields.Item("Lotes").Value));

                                            lin = new lineas_Factura()
                                            {
                                                codigo = Convert.ToString(oRecordSetLin.Fields.Item("ItemCode").Value),
                                                cantidad = Convert.ToInt32(oRecordSetLin.Fields.Item("Quantity").Value),
                                                serie = s,
                                                lote = l
                                            };
                                            oRecordSetLin.MoveNext();
                                        }
                                        else if (!string.IsNullOrEmpty(lin.codigo))
                                        {
                                            trasladoDev.Add(lin);
                                            numLinea++;
                                            lin = new lineas_Factura();
                                            s = new Dictionary<string, string>();
                                            l = new List<string>();
                                        }
                                        else
                                        {
                                            numLinea++;
                                        }
                                        if (oRecordSetLin.EoF)
                                        {
                                            trasladoDev.Add(lin);
                                            s = new Dictionary<string, string>();
                                            l = new List<string>();
                                        }
                                       
                                    }
                                    // escribirLog("FacturarTraslado: Almaceno artículos del traslado Devolución");
                                    //bool validar = false;
                                    //foreach(var T in traslado)
                                    //{
                                    //    Console.WriteLine("Codigo: " + T.codigo + "\ncantidad: " + T.cantidad);
                                    //    foreach (var ser in T.serie)
                                    //    {
                                    //        Console.WriteLine("Serie: " + ser.Key + "\nValue: " + ser.Value);
                                    //    }
                                    //}
                                    //Console.WriteLine("----------------DEVOLUCION---------------");
                                    //foreach (var T in trasladoDev)
                                    //{
                                    //    Console.WriteLine("Codigo: " + T.codigo + "\ncantidad: " + T.cantidad);
                                    //    foreach (var ser in T.serie)
                                    //    {
                                    //        Console.WriteLine("Serie: " + ser.Key + "\nValue: " + ser.Value);
                                    //    }
                                    //}
                                    //Comparación de los traslados 
                                    List<string> keys = new List<string>();
                                    foreach (var articuloTras in traslado)
                                    {
                                        foreach (var articuloDev in trasladoDev)
                                        {
                                            if (articuloTras.codigo == articuloDev.codigo)
                                            {
                                                foreach (var serDev in articuloDev.serie)
                                                {
                                                    if (articuloTras.serie.ContainsKey(serDev.Key)) keys.Add(serDev.Key);//tras.serie.Remove(serDev.Key);
                                                }
                                            }

                                        }
                                    }
                                    //Console.WriteLine("------------Eliminar series----------");
                                    //foreach (var k in keys)
                                    //{
                                    //    Console.WriteLine(k);
                                    //}
                                    //Elimina las series y articulos devueltos en el traslado relacionado
                                    int contador = 0;
                                    while (contador <= traslado.Count - 1)
                                    {
                                        foreach (var key in keys)
                                        {
                                            if (traslado[contador].serie.ContainsKey(key))
                                            {
                                                traslado[contador].serie.Remove(key);
                                                traslado[contador].cantidad -= 1;
                                            }
                                            if (traslado[contador].serie.Count == 0 && traslado[contador].cantidad == 0)
                                            {
                                                traslado.Remove(traslado[contador]);
                                            }
                                        }
                                        contador++;
                                    }
                                    keys.Clear();
                                    dev = true;
                                    //Console.WriteLine("----------------TRASLADO FINAL--------------");
                                    //foreach (var T in traslado)
                                    //{
                                    //    Console.WriteLine("Codigo: " + T.codigo + "\ncantidad: " + T.cantidad);
                                    //    foreach (var ser in T.serie)
                                    //    {
                                    //        Console.WriteLine("Serie: " + ser.Key + "\nValue: " + ser.Value);
                                    //    }
                                    //}
                                    //escribirLog("FacturarTraslado: Comparación Traslados");
                                }
                                else
                                {
                                    SBO_Application.StatusBar.SetText("Es necesario asignar la transferencia de Stock como devolución", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    Business.escribirLog("FacturarTransferenciaStock: Es necesario asignar la transferencia de Stock como devolución");
                                    return;
                                }
                            }
                            catch (Exception ex)
                            {
                                escribirLog("TrasladoDevolución: "+ ex.Message);
                                SBO_Application.StatusBar.SetText("Ocurrio un problema: " + "TrasladoDevolución: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                return;
                            }
                            finally
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSetLin);
                            }
                        }
                        else
                        {

                            try
                            {
                                lineas_Factura lin = new lineas_Factura();
                                Dictionary<string, string> s = new Dictionary<string, string>();
                                List<string> l = new List<string>();

                                sSQL = string.Format(Consultas.Default.TransferenciaCab, docEntryDev);
                                oRecordSet.DoQuery(sSQL);

                                cab = new Factura_Traslado();
                                cab.SN = SNFactura;
                                cab.fechaPed = Convert.ToString(oRecordSet.Fields.Item("Fecha Ped.").Value);
                                cab.fechaCir = Convert.ToString(oRecordSet.Fields.Item("Fecha Cir.").Value);
                                cab.nombreMed = Convert.ToString(oRecordSet.Fields.Item("Nombre Med.").Value);
                                cab.nombrePac = Convert.ToString(oRecordSet.Fields.Item("Nombre Pac.").Value);
                                cab.clinica = Convert.ToString(oRecordSet.Fields.Item("Clinica").Value);
                                cab.cedula = Convert.ToString(oRecordSet.Fields.Item("Cedula").Value);
                                cab.aFactura = Convert.ToString(oRecordSet.Fields.Item("Se factura.").Value);
                                cab.ubicacion = Convert.ToString(oRecordSet.Fields.Item("AbsEntry").Value);
                                cab.tipoDoc = Convert.ToString(oRecordSet.Fields.Item("Documento").Value);
                                cab.transRef = docEntryDev;
                                // Datos Lineas
                                sSQL = string.Format(Consultas.Default.Articulos_Series, docEntryDev);
                                oRecordSetLin.DoQuery(sSQL);
                                int numLinea = Convert.ToInt32(oRecordSetLin.Fields.Item("LineNum").Value);

                                s = new Dictionary<string, string>();
                                l = new List<string>();
                                oRecordSetLin.MoveFirst();
                                while (!oRecordSetLin.EoF)
                                {
                                    if (Convert.ToInt32(oRecordSetLin.Fields.Item("LineNum").Value) == numLinea)
                                    {
                                        s.Add(Convert.ToString(oRecordSetLin.Fields.Item("Series").Value), Convert.ToString(oRecordSetLin.Fields.Item("Series").Value));
                                        l.Add(Convert.ToString(oRecordSetLin.Fields.Item("Lotes").Value));

                                        lin = new lineas_Factura()
                                        {
                                            codigo = Convert.ToString(oRecordSetLin.Fields.Item("ItemCode").Value),
                                            cantidad = Convert.ToInt32(oRecordSetLin.Fields.Item("Quantity").Value),
                                            almacen = Convert.ToString(oRecordSetLin.Fields.Item("WhsCode").Value),
                                            ciudad = Convert.ToString(oRecordSetLin.Fields.Item("OcrCode").Value),
                                            centroCostos = Convert.ToString(oRecordSetLin.Fields.Item("OcrCode2").Value),
                                            serie = s,
                                            lote = l
                                        };
                                        oRecordSetLin.MoveNext();
                                    }
                                    else if (!string.IsNullOrEmpty(lin.codigo))
                                    {
                                        traslado.Add(lin);
                                        numLinea++;
                                        lin = new lineas_Factura();
                                        s = new Dictionary<string, string>();
                                        l = new List<string>();
                                    }
                                    else
                                    {
                                        numLinea++;
                                    }
                                    if (oRecordSetLin.EoF)
                                    {
                                        traslado.Add(lin);
                                        s = new Dictionary<string, string>();
                                        l = new List<string>();
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                escribirLog("TrasladoSinDev: "+ex.Message);
                                SBO_Application.StatusBar.SetText("TrasladoSinDev: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                return;
                            }
                            finally
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSetLin);
                            }

                        }
                        //Recibe los articulos a facturar
                        FacturarLin = traslado;
                        CrearBorradorFV(cab, FacturarLin, dev);
                        GC.Collect();
                    }
                    else
                    {
                        SBO_Application.StatusBar.SetText("Es necesario diligenciar el campo SN Factura", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        Business.escribirLog("FacturarTransferenciaStock: Es necesario diligenciar el campo SN Factura");
                        GC.Collect();
                        return;
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
                Business.escribirLog("FacturarTraslado: " + ex.Message);
                return;
            }
        }

        //Crear documento prelimninar de la factura por medio de la SL
        public void CrearBorradorFV(Factura_Traslado facturarCab, List<lineas_Factura> facturarLin, bool dev)
        {
            try
            {
                List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
                Dictionary<string, object> row;
                foreach (var linea in facturarLin)
                {
                    row = new Dictionary<string, object>();
                    row.Add("ItemCode", linea.codigo); //Codigo articulo
                    row.Add("Quantity", linea.cantidad);//Cantidad
                    row.Add("WarehouseCode", linea.almacen); //Almacen     
                    row.Add("CostingCode", linea.ciudad); //Ciudad   
                    row.Add("CostingCode2", linea.centroCostos); //Centro de costos   
                    rows.Add(row);
                }
                JObject jsonObj = new JObject();
                jsonObj.Add("CardCode", facturarCab.SN);
                jsonObj.Add("DocObjectCode", "13");// Borrar documento Factura de venta
                jsonObj.Add("U_SCL_Fecha01", facturarCab.fechaPed);
                jsonObj.Add("U_SCL_Fecha02", facturarCab.fechaCir);
                jsonObj.Add("U_SCL_Nombre", facturarCab.nombreMed);
                jsonObj.Add("U_SCL_Nombre01", facturarCab.nombrePac);
                jsonObj.Add("U_SCL_Clinica", facturarCab.clinica);
                jsonObj.Add("U_SCL_Nombre02", facturarCab.aFactura);
                jsonObj.Add("U_SCL_Numero", facturarCab.cedula);
                jsonObj.Add("U_SCL_TipoDoc", facturarCab.tipoDoc);
                //                jsonObj.Add("U_SCL_TransRef", facturarCab.transRef);
                if (dev.Equals(true))
                {
                    jsonObj.Add("U_SCL_NroTraslado", facturarCab.transOrg);
                }
                else
                {
                    jsonObj.Add("U_SCL_NroTraslado", facturarCab.transRef);
                }
                jsonObj.Add("DocumentLines", JArray.FromObject(rows));
                var body = JsonConvert.SerializeObject(jsonObj);
                string sessionID = oConnection.ConexionServiceLayer();
                RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
                RestRequest request = new RestRequest("Drafts", Method.POST);
                request.RequestFormat = DataFormat.Json;
                request.AddCookie("B1SESSION", sessionID);
                request.AddParameter("application/json", body, ParameterType.RequestBody);
                RestResponse response = (RestResponse)cliente.Execute(request);
                var res = response.Content;
                //Console.WriteLine(response.StatusDescription);
                if (response.StatusCode.Equals(HttpStatusCode.Created))
                {
                    dynamic dynJson = JsonConvert.DeserializeObject(res);
                    string numeroDoc = null;
                    string docEntry = null;
                    foreach (var item in dynJson)
                    {
                        if (item.Name == "DocNum") numeroDoc = item.Value;
                        if (item.Name == "DocEntry") docEntry = item.Value;
                    }
                    SBO_Application.MessageBox("Creación del documento preliminar número " + numeroDoc + " satisfactoria! ");
                    escribirLog("CrearBorradorFactura: Creación del documento preliminar número " + numeroDoc + " satisfactoria!");
                    ActualizarSeries(facturarLin, facturarCab.ubicacion, sessionID, docEntry);
                    ActualizarTraslado(facturarCab.transRef, docEntry, sessionID);
                }
                else
                {
                    var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                    var jvalue = (JValue)jobject["error"]["message"]["value"];
                    SBO_Application.SetStatusBarMessage(jvalue.Value + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    escribirLog("CrearBorradorFactura: " + jvalue.Value + oCompany.GetLastErrorDescription());
                }
            }
            catch (Exception ex)
            {
                //SBO_Application.MessageBox(ex.Message);
                SBO_Application.StatusBar.SetText("CrearBorrador" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                Business.escribirLog("CrearBorrador: " + ex.Message);
                return;
            }

        }

        //Actualizar series y ubicaciones del documento preliminar creado
        public void ActualizarSeries(List<lineas_Factura> facturarLin, string ubi, string sessionID, string docEntry)
        {
            //var estructura2 = "{*ItemCode*:*?*,*SerialNumbers*:%}";
            string result = "";
            //estructura2 = (estructura2.ToString().Replace('*', '"').Trim());
            Dictionary<string, string> values = new Dictionary<string, string>();
            Dictionary<string, object> row = new Dictionary<string, object>();
            List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
            Dictionary<string, string> rowUbi = new Dictionary<string, string>();
            List<Dictionary<string, string>> rowsUbi = new List<Dictionary<string, string>>();
            try
            {
                foreach (var linea in facturarLin)
                {
                    rows = new List<Dictionary<string, object>>();
                    rowsUbi = new List<Dictionary<string, string>>();
                    //var estructura2 = "{*ItemCode*:*?*,*SerialNumbers*:%}";
                    var estructura2 = ubi == "0" ? "{*ItemCode*:*?*,*SerialNumbers*:%}" : "{*ItemCode*:*?*,*SerialNumbers*:%,*DocumentLinesBinAllocations*:@}"; //"BinAbsEntry": 82,
                    estructura2 = (estructura2.ToString().Replace('*', '"').Trim());
                    estructura2 = estructura2.Replace("?", linea.codigo);
                    foreach (var lin in linea.serie)
                    {
                        row = new Dictionary<string, object>();
                        row.Add("InternalSerialNumber", lin.Value); //Codigo articulo
                        rows.Add(row);

                        rowUbi = new Dictionary<string, string>();
                        rowUbi.Add("BinAbsEntry", ubi); //Codigo articulo
                        rowsUbi.Add(rowUbi);
                    }
                    string json = JsonConvert.SerializeObject(rows);
                    estructura2 = estructura2.Replace("%", json);
                    if (ubi != "0") json = JsonConvert.SerializeObject(rowsUbi); estructura2 = estructura2.Replace("@", json);
                    result += estructura2;
                    if (linea.serie.Count > 0) result += ",";
                }
                var estructura = "{*DocumentLines*:[¿]}"; //"BinAbsEntry": 82,
                var body = (estructura.ToString().Replace('*', '"').Trim());
                estructura = body.Replace("¿", result);

                //string sessionID = oConnection.ConexionServiceLayer();
                RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
                RestRequest request = new RestRequest("Drafts(" + docEntry + ")", Method.PATCH);
                request.RequestFormat = DataFormat.Json;
                request.AddCookie("B1SESSION", sessionID);
                request.AddParameter("application/json", estructura, ParameterType.RequestBody);
                RestResponse response = (RestResponse)cliente.Execute(request);
                var res = response.Content;
                if (response.StatusCode.Equals(HttpStatusCode.Created) || response.StatusCode.Equals(HttpStatusCode.NoContent))
                {
                    //SBO_Application.MessageBox("Se actualizaron las series del documento " + docEntry + " satisfactoriamente! ");
                    escribirLog("Se actualizaron las series del documento (DocEntry) " + docEntry);
                }
                else
                {
                    var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                    var jvalue = (JValue)jobject["error"]["message"]["value"];
                    SBO_Application.SetStatusBarMessage(jvalue.Value + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    escribirLog("ActualizarSeriesDocPrel: " + jvalue.Value + oCompany.GetLastErrorDescription());
                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText("ActualizarSeriesDocPrel" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                Business.escribirLog("ActualizarSeriesDocPrel: " + ex.Message);                
            }
        }
        //Asinagnar el docEntry del documento preliminar creado
        public void ActualizarTraslado(string docEntryTras, string docEntryFac, string sessionID)
        {
            try
            {
                Dictionary<string, object> row = new Dictionary<string, object>(){
                    { "U_SCL_DocPrel", docEntryFac }//Codigo articulo
                };
                var body = JsonConvert.SerializeObject(row);
                RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
                RestRequest request = new RestRequest("StockTransfers(" + docEntryTras + ")", Method.PATCH);
                request.RequestFormat = DataFormat.Json;
                request.AddCookie("B1SESSION", sessionID);
                request.AddParameter("application/json", body, ParameterType.RequestBody);
                RestResponse response = (RestResponse)cliente.Execute(request);
                var res = response.Content;
                if (response.StatusCode.Equals(HttpStatusCode.Created) || response.StatusCode.Equals(HttpStatusCode.NoContent))
                {
                    //SBO_Application.MessageBox("Se actualizaron las series del documento " + docEntry + " satisfactoriamente! ");
                    // escribirLog("Actualizar Documento preliminar " + docEntryFac);
                }
                else
                {
                    var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                    var jvalue = (JValue)jobject["error"]["message"]["value"];
                    SBO_Application.SetStatusBarMessage(jvalue.Value + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    escribirLog("ActualizacionCampoTraslado: " + jvalue.Value + oCompany.GetLastErrorDescription());
                }
            }
            catch (Exception ex)
            {
                Business.escribirLog("ActualizarTraslado: " + ex.Message);
            }
        }
        #endregion



        private string verinit(string cardCode)
        {
            string Retorno = "";
            string nit = (cardCode.Substring(0, cardCode.Length - 2));
            int num;
            int sumaTotal = 0;
            int codigo = 0;
            int[] numerosPrimos = { 3, 7, 13, 17, 19, 23, 29, 37, 41, 43, 47, 53, 59, 67, 71 };
            nit = InvertirString(nit);
            for (int i = 0; i < nit.Length; i++)
            {
                num = Int32.Parse(nit.Substring(i, 1));
                sumaTotal += (num * numerosPrimos[i]);
            }
            codigo = sumaTotal % 11;
            if (codigo > 1)
                codigo = 11 - codigo;
            if (codigo > 10)
                codigo = 0;

            Retorno = Convert.ToString(codigo);
            return Retorno;
        }

        private string InvertirString(string nit)
        {
            char[] arr = nit.ToCharArray();
            Array.Reverse(arr);
            return new string(arr);
        }

        /// <summary>
        /// Manaejo de eventos retencion articulos
        /// </summary>
        /// <param name="pVal">Tipo de evento</param>
        /// <param name="formUID">Identificador del formulario</param>
        private void EventFieldRet(ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.Item oItem = null;
            // oForm = SBO_Application.Forms.Item(FormUID);

            if (pVal.FormUID == "AsisBalTer" && pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
            {
                try
                {
                    oForm = SBO_Application.Forms.Item(pVal.FormUID);
                    BalanceTer balance = new BalanceTer(oCompany, SBO_Application);
                    balance.agregarComponentesForm();
                    balance.cargarDatos();
                    oForm.Visible = true;

                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("LoadItem: " + ex.Message);
                }
            }
            else if (pVal.FormUID == "AsisReclas")
            {
                //if (pVal.FormUID == "AsisReclas" && pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                //{
                //    try
                //    {
                //        asistente.cargarDatos();
                //        //oForm.Visible = true;//
                //    }
                //    catch (Exception ex)
                //    {
                //        SBO_Application.MessageBox(ex.Message);
                //        //Business.escribirLog("LoadItem: " + ex.Message);
                //    }
                //}
                //                else
                if (pVal.FormUID == "AsisReclas" && pVal.EventType == BoEventTypes.et_FORM_CLOSE && pVal.Before_Action == true)
                {
                    try
                    {
                        asistente.Cerrar_Ventana();
                    }
                    catch (Exception ex)
                    {
                        SBO_Application.MessageBox(ex.Message);
                        //Business.escribirLog("LoadItem: " + ex.Message);
                    }
                }
                else if (pVal.FormUID == "AsisReclas" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "btnAsiento" && pVal.BeforeAction == true)
                {
                    try
                    {
                        asistente.Click_btnAsiento();
                    }
                    catch (Exception ex)
                    {
                        SBO_Application.MessageBox(ex.Message);
                        Business.escribirLog("LoadItem: " + ex.Message);
                    }
                }
                else if (pVal.FormUID == "AsisReclas" && (pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "btnAgRet") && pVal.Before_Action == true)
                {
                    try
                    {
                        asistente.Click_btnAgregarRet();

                    }
                    catch (Exception ex)
                    {
                        SBO_Application.MessageBox(ex.Message);
                        Business.escribirLog("LoadItem: " + ex.Message);
                    }
                }
                else if (pVal.FormUID == "AsisReclas" && (pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "btnAñadir") && pVal.Before_Action == true)
                {
                    try
                    {
                        asistente.Click_btnAñadir();

                    }
                    catch (Exception ex)
                    {
                        SBO_Application.MessageBox(ex.Message);
                        Business.escribirLog("LoadItem: " + ex.Message);
                    }
                }
                else if (pVal.FormUID == "AsisReclas" && (pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "btnElim") && pVal.Before_Action == true)
                {
                    try
                    {
                        asistente.Click_btnEliminar();

                    }
                    catch (Exception ex)
                    {
                        SBO_Application.MessageBox(ex.Message);
                        Business.escribirLog("LoadItem: " + ex.Message);
                    }
                }
                else if (pVal.FormUID == "AsisReclas" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "btnCancel" && pVal.BeforeAction == true)
                {
                    try
                    {
                        asistente.Click_btnCancel();
                    }
                    catch (Exception ex)
                    {
                        SBO_Application.MessageBox(ex.Message);
                        Business.escribirLog("LoadItem: " + ex.Message);
                    }
                }
            }
            if (pVal.FormUID == "RETITEM" && pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
            {
                try
                {
                    oForm = SBO_Application.Forms.Item(pVal.FormUID);

                    CrearItemsRet(pVal.FormUID);

                    oItem = oForm.Items.Item("3");
                    oMatrix = (Matrix)oItem.Specific;

                    Recordset oRecordSet = null;
                    oRecordSet = ((Recordset)(oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)));
                    oRecordSet.DoQuery(Consultas.Default.ListarRet);

                    int i = 1;
                    while (oRecordSet.EoF == false)
                    {
                        string sSQL = "";
                        string CodRet = "";
                        CodRet = Convert.ToString(oRecordSet.Fields.Item(0).Value);
                        oMatrix.AddRow();
                        ((EditText)oMatrix.Columns.Item(0).Cells.Item(i).Specific).Value = (i).ToString();
                        ((EditText)oMatrix.Columns.Item(1).Cells.Item(i).Specific).Value = Convert.ToString(oRecordSet.Fields.Item(0).Value);
                        ((EditText)oMatrix.Columns.Item(2).Cells.Item(i).Specific).Value = Convert.ToString(oRecordSet.Fields.Item(1).Value);

                        Recordset oRS = null;
                        oRS = ((Recordset)(oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)));
                        // oRS.DoQuery("Select * From \"@SCL_ITM4\" Where \"Code\" = '" + CodigoArticulo + "' And \"Object\" = '" + CodRet + "'");

                        sSQL = string.Format(Consultas.Default.CulculoRet, CodigoArticulo, CodRet);
                        escribirLog("ConsultaRet: " + sSQL);
                        oRS.DoQuery(sSQL);


                        if (oRS.RecordCount > 0)
                        {
                            ((CheckBox)oMatrix.Columns.Item(3).Cells.Item(i).Specific).Checked = true;
                        }

                        oRecordSet.MoveNext();
                        i = i + 1;
                    }
                    oForm.Mode = BoFormMode.fm_OK_MODE;

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                    oRecordSet = null;
                    GC.Collect();
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("LoadItem: " + ex.Message);
                }
            }
            if (pVal.FormUID == "RETITEM" && pVal.FormMode == 2 && pVal.ItemUID == "1" && pVal.EventType == BoEventTypes.et_CLICK && pVal.Before_Action == false)
            {
                try
                {
                    string ItemCode = "";
                    oForm = SBO_Application.Forms.Item(pVal.FormUID);

                    oItem = oForm.Items.Item("ItemCode");
                    oEdit = oItem.Specific;
                    ItemCode = oEdit.Value;

                    Recordset oRecordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    oForm = SBO_Application.Forms.Item(pVal.FormUID);

                    oItem = oForm.Items.Item("3");
                    oMatrix = oItem.Specific;

                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {
                        string CodRet = "";
                        string sSQL = "";
                        string sSQL1 = "";
                        bool Retselect = false;
                        CodRet = ((SAPbouiCOM.EditText)oMatrix.GetCellSpecific("3", i)).Value;
                        Retselect = ((SAPbouiCOM.CheckBox)oMatrix.GetCellSpecific("2", i)).Checked;

                        //oRecordset.DoQuery("Select * From \"@SCL_ITM4\" Where \"Code\" = '" + ItemCode + "' And \"Object\" = '" + CodRet + "'");

                        sSQL1 = string.Format(Consultas.Default.CalculoRetItem, ItemCode, CodRet);
                        oRecordset.DoQuery(sSQL1);

                        if (Retselect == true && oRecordset.RecordCount <= 0)

                        {
                            sSQL = string.Format(Consultas.Default.AddRetItem, ItemCode, CodRet);
                            oRecordset.DoQuery(sSQL);
                        }
                        else if (Retselect == false && oRecordset.RecordCount > 0)
                        {
                            sSQL = string.Format(Consultas.Default.DelRetItem, ItemCode, CodRet);
                            oRecordset.DoQuery(sSQL);
                        }
                    }

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
                    oRecordset = null;
                    GC.Collect();
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("LoadItem: " + ex.Message);
                }
            }
            if (pVal.FormUID == "AUTORET" && pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
            {
                try
                {
                    oForm = SBO_Application.Forms.Item(pVal.FormUID);

                    CrearItemsRetSN(pVal.FormUID);

                    oItem = oForm.Items.Item("3");
                    oMatrix = (Matrix)oItem.Specific;

                    Recordset oRecordSet = null;
                    oRecordSet = ((Recordset)(oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)));
                    oRecordSet.DoQuery(Consultas.Default.ListarAutoRet);

                    int i = 1;
                    while (oRecordSet.EoF == false)
                    {
                        string CodRet = "";
                        CodRet = Convert.ToString(oRecordSet.Fields.Item(0).Value);
                        oMatrix.AddRow();
                        ((EditText)oMatrix.Columns.Item(0).Cells.Item(i).Specific).Value = (i).ToString();
                        ((EditText)oMatrix.Columns.Item(1).Cells.Item(i).Specific).Value = Convert.ToString(oRecordSet.Fields.Item(0).Value);
                        ((EditText)oMatrix.Columns.Item(2).Cells.Item(i).Specific).Value = Convert.ToString(oRecordSet.Fields.Item(1).Value);

                        Recordset oRS = null;
                        oRS = ((Recordset)(oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)));
                        oRS.DoQuery("Select * From \"@SCL_CRD4\" Where \"Code\" = '" + CodigoSocio + "' And \"Object\" = '" + CodRet + "'");

                        if (oRS.RecordCount > 0)
                        {
                            ((CheckBox)oMatrix.Columns.Item(3).Cells.Item(i).Specific).Checked = true;
                        }

                        oRecordSet.MoveNext();
                        i = i + 1;
                    }
                    oForm.Mode = BoFormMode.fm_OK_MODE;

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                    oRecordSet = null;
                    GC.Collect();
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("LoadItem: " + ex.Message);
                }
            }

            if (pVal.FormUID == "AUTORET" && pVal.FormMode == 2 && pVal.ItemUID == "1" && pVal.EventType == BoEventTypes.et_CLICK && pVal.Before_Action == false)
            {
                try
                {
                    string CardCode = "";
                    oForm = SBO_Application.Forms.Item(pVal.FormUID);

                    oItem = oForm.Items.Item("CardCode");
                    oEdit = oItem.Specific;
                    CardCode = oEdit.Value;

                    Recordset oRecordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    oForm = SBO_Application.Forms.Item(pVal.FormUID);

                    oItem = oForm.Items.Item("3");
                    oMatrix = oItem.Specific;

                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {
                        string CodRet = "";
                        string sSQL = "";
                        bool Retselect = false;
                        CodRet = ((SAPbouiCOM.EditText)oMatrix.GetCellSpecific("3", i)).Value;
                        Retselect = ((SAPbouiCOM.CheckBox)oMatrix.GetCellSpecific("2", i)).Checked;

                        oRecordset.DoQuery("Select * From \"@SCL_CRD4\" Where \"Code\" = '" + CardCode + "' And \"Object\" = '" + CodRet + "'");
                        if (Retselect == true && oRecordset.RecordCount <= 0)
                        {
                            sSQL = string.Format(Consultas.Default.AddAutoRet, CardCode, CodRet);
                            oRecordset.DoQuery(sSQL);
                        }
                        else if (Retselect == false && oRecordset.RecordCount > 0)
                        {
                            sSQL = string.Format(Consultas.Default.DelAutoRet, CardCode, CodRet);
                            oRecordset.DoQuery(sSQL);
                        }
                    }

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
                    oRecordset = null;
                    GC.Collect();
                }
                catch (Exception ex)
                {
                    SBO_Application.MessageBox(ex.Message);
                    Business.escribirLog("LoadItem: " + ex.Message);
                }
            }


            //if (pVal.FormUID == "CIERREFISCAL01" && pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
            //{
            //    try
            //    {
            //        oForm = SBO_Application.Forms.Item(pVal.FormUID);
            //        CrearItemsCierre(pVal.FormUID);
            //    }
            //    catch (Exception ex)
            //    {
            //        SBO_Application.MessageBox(ex.Message);
            //        Business.escribirLog("LoadItem: " + ex.Message);
            //    }
            //}

            //if (pVal.FormUID == "CIERREFISCAL01" && pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST /*&& pVal.Before_Action == true*/)
            //{
            //    try
            //    {
            //        SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
            //        oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
            //        string sCFL_ID = null;
            //        sCFL_ID = oCFLEvento.ChooseFromListUID;
            //        SAPbouiCOM.Form oForm = null;
            //        oForm = SBO_Application.Forms.Item(FormUID);
            //        SAPbouiCOM.ChooseFromList oCFL = null;
            //        oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
            //        if (oCFLEvento.BeforeAction == false)
            //        {
            //            SAPbouiCOM.DataTable oDataTable = null;
            //            oDataTable = oCFLEvento.SelectedObjects;
            //            string val = null;
            //            try
            //            {
            //                val = System.Convert.ToString(oDataTable.GetValue(0, 0));
            //            }
            //            catch (Exception ex)
            //            {

            //            }
            //            switch (pVal.ItemUID)
            //            {
            //                case "ToAcct":
            //                    oForm.DataSources.UserDataSources.Item("EditToC").ValueEx = val;
            //                    break;
            //                case "FromAcct":
            //                    oForm.DataSources.UserDataSources.Item("EditFromC").ValueEx = val;
            //                    break;
            //                case "ToSocio":
            //                    oForm.DataSources.UserDataSources.Item("EditToS").ValueEx = val;
            //                    break;
            //                case "FromSocio":
            //                    oForm.DataSources.UserDataSources.Item("EditFromS").ValueEx = val;
            //                    break;
            //            }
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        SBO_Application.MessageBox(ex.Message);
            //        Business.escribirLog("LoadItem: " + ex.Message);
            //    }
            //}

            //if (pVal.FormUID == "CIERREFISCAL01" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "4" && pVal.BeforeAction == true)
            //{
            //    try
            //    {
            //        string fechaini = "";
            //        string fechafin = "";
            //        string ToAcct = "";
            //        string FromAcct = "";
            //        string ToCardCode = "";
            //        string FromCardCode = "";
            //        string TransCode = "";
            //        string Year = "";

            //        oForm = SBO_Application.Forms.Item(FormUID);

            //        oForm.PaneLevel = 2;

            //        oItem = oForm.Items.Item("FechaIni");
            //        oEdit = (SAPbouiCOM.EditText)oItem.Specific;
            //        fechaini = oEdit.Value;

            //        oItem = oForm.Items.Item("FechaFin");
            //        oEdit = (SAPbouiCOM.EditText)oItem.Specific;
            //        fechafin = oEdit.Value;

            //        oItem = oForm.Items.Item("ToAcct");
            //        oEdit = (SAPbouiCOM.EditText)oItem.Specific;
            //        ToAcct = oEdit.Value;

            //        oItem = oForm.Items.Item("FromAcct");
            //        oEdit = (SAPbouiCOM.EditText)oItem.Specific;
            //        FromAcct = oEdit.Value;

            //        oItem = oForm.Items.Item("ToSocio");
            //        oEdit = (SAPbouiCOM.EditText)oItem.Specific;
            //        ToCardCode = oEdit.Value;

            //        oItem = oForm.Items.Item("FromSocio");
            //        oEdit = (SAPbouiCOM.EditText)oItem.Specific;
            //        FromCardCode = oEdit.Value;

            //        oItem = oForm.Items.Item("TransCode");
            //        oComboBox = (SAPbouiCOM.ComboBox)oItem.Specific;
            //        TransCode = oComboBox.Value;

            //        oItem = oForm.Items.Item("Year");
            //        oComboBox = (SAPbouiCOM.ComboBox)oItem.Specific;
            //        Year = oComboBox.Value;

            //        if (!string.IsNullOrEmpty(fechaini) && !string.IsNullOrEmpty(fechafin))
            //        {
            //            cargarGrid("CIERREFISCAL01", fechaini, fechafin);
            //        }
            //        else
            //        {
            //            SBO_Application.MessageBox("Debe diligenciar los parametros de fechas");
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        SBO_Application.MessageBox(ex.Message);
            //        Business.escribirLog("LoadItem: " + ex.Message);
            //    }
            //}

            //if (pVal.FormUID == "CIERREFISCAL01" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "3" && pVal.BeforeAction == true)
            //{
            //    try
            //    {
            //        oForm = SBO_Application.Forms.Item(FormUID);

            //        oForm.PaneLevel = 1;
            //    }
            //    catch (Exception ex)
            //    {
            //        SBO_Application.MessageBox(ex.Message);
            //        Business.escribirLog("LoadItem: " + ex.Message);
            //    }
            //}

            //if (pVal.FormUID == "CIERREFISCAL01" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "129" && pVal.BeforeAction == true)
            //{
            //    try
            //    {
            //        SAPbobsCOM.JournalEntries oObjJournal = null;

            //        oForm = SBO_Application.Forms.Item(FormUID);

            //        oItem = oForm.Items.Item("Grid");
            //        oGrid = oItem.Specific;

            //        oObjJournal = (SAPbobsCOM.JournalEntries)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

            //        oObjJournal.ReferenceDate = DateTime.Now;
            //        oObjJournal.TransactionCode = "CIFI";


            //        int Count = oGrid.Rows.Count;
            //        double valorCredit = 0;

            //        for (int i = 0; i <= oGrid.Rows.Count - 1; i++)
            //        {
            //            string addValor = oGrid.DataTable.GetValue("Incluir", i);
            //            if (addValor == "Y")
            //            {
            //                oObjJournal.Lines.AccountCode = oGrid.DataTable.GetValue("DebPayAcct", i);
            //                oObjJournal.Lines.Debit = oGrid.DataTable.GetValue("Balance", i);

            //                valorCredit = valorCredit + oGrid.DataTable.GetValue("Balance", i);
            //                if (i != Count)
            //                {
            //                    oObjJournal.Lines.Add();
            //                }
            //            }
            //        }

            //        //oObjJournal.Lines.Add();
            //        oObjJournal.Lines.AccountCode = "42352505";//oGrid.DataTable.GetValue("DebPayAcct", 1);
            //        oObjJournal.Lines.Credit = valorCredit;

            //        lRetCode = oObjJournal.Add();

            //        if (lRetCode != 0)
            //        {
            //            oCompany.GetLastError(out lRetCode, out sErrMsg);
            //            SBO_Application.MessageBox("Error: " + sErrMsg);
            //        }
            //        else
            //        {
            //            string DocEntry = "";
            //            oCompany.GetNewObjectCode(out DocEntry);
            //            SBO_Application.MessageBox("Asiento creado: " + DocEntry);
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        SBO_Application.MessageBox(ex.Message);
            //        Business.escribirLog("LoadItem: " + ex.Message);
            //    }
            //}
        }

        private void EventFieldParamLoc(ItemEvent pVal, string FormUID)
        {
            try
            {
                oForm = SBO_Application.Forms.Item(pVal.FormUID);
                //if (pVal.FormTypeEx == "SCL_ParamIniLoc" && pVal.EventType == BoEventTypes.et_FORM_LOAD  && pVal.Before_Action == false)
                //{
                //    try
                //    {
                //        oItem = oForm.Items.Item("fldAS");
                //        oItem.Click();
                //    }
                //    catch (Exception ex)
                //    {
                //        SBO_Application.MessageBox(ex.Message);
                //    }
                //}
                if (pVal.FormTypeEx == "SCL_ParamIniLoc" && (pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "btnActSN") && pVal.Before_Action == true)
                {
                    try
                    {
                        oItem = oForm.Items.Item("btnActSN");
                        oItem.Enabled = false;
                        AsignarTerceroAsientos();
                        oItem.Enabled = true;
                    }
                    catch (Exception ex)
                    {
                        SBO_Application.MessageBox(ex.Message);
                        escribirLog("btnActSN: " + ex.Message);
                    }
                }
                if (pVal.FormTypeEx == "SCL_ParamIniLoc" && (pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "btnActCont") && pVal.Before_Action == true)
                {
                    try
                    {
                        oItem = oForm.Items.Item("btnActCont");
                        oItem.Enabled = false;
                        ActualizarTipoContAsientos();
                        oItem.Enabled = true;
                    }
                    catch (Exception ex)
                    {
                        SBO_Application.MessageBox(ex.Message);
                        escribirLog("btnActCont: " + ex.Message);
                    }
                }
                if (pVal.FormTypeEx == "SCL_ParamIniLoc" && (pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "btnAnulNC") && pVal.Before_Action == true)
                {
                    try
                    {
                        oItem = oForm.Items.Item("btnAnulNC");
                        oItem.Enabled = false;
                        AnularNCVentas();
                        AnularNCCompras();
                        oItem.Enabled = true;
                    }
                    catch (Exception ex)
                    {
                        SBO_Application.MessageBox(ex.Message);
                        escribirLog("btnAnulNC: " + ex.Message);
                    }
                }
//cuando pulzo actualizar los check para el proceso del batch 
                if (pVal.FormTypeEx == "SCL_ParamIniLoc" && (pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "1") && pVal.Before_Action == true)
                {
                    string AutRet = string.Empty;
                    string AgrupRet = string.Empty;
                    string AsRet = string.Empty;
                    string sqlUpdate = string.Empty;
                    Recordset oRecordset2 = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    SAPbobsCOM.Recordset oRecUpdate = null;
                    oRecUpdate = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                    SAPbouiCOM.CheckBox oCheckBox1;
                    oCheckBox1 = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkAutB").Specific;
                    SAPbouiCOM.CheckBox oCheckBox2;
                    oCheckBox2 = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkAsRes").Specific;
                    SAPbouiCOM.CheckBox oCheckBox3;
                    oCheckBox3 = (SAPbouiCOM.CheckBox)oForm.Items.Item("chkAgrup").Specific;


                    if (oCheckBox1.Checked) { AutRet = "Y"; } else { AutRet = "N"; }
                    if (oCheckBox2.Checked) { AsRet = "Y"; } else { AsRet = "N"; }
                    if (oCheckBox3.Checked) { AgrupRet = "Y"; } else { AgrupRet = "N"; }
                    try
                    {
                        sqlUpdate = "UPDATE \"@SCL_LOC_VERSION\" SET \"U_SCL_Agrup\"='" + AgrupRet + "',\"U_SCL_AsRes\"='" + AsRet + "',\"U_SCL_AutB\"='" + AutRet + "'";
                        oRecordset2.DoQuery(sqlUpdate);
                        //bool ban=true;
                    }
                    catch (Exception ex)
                    {
                        SBO_Application.MessageBox(ex.Message);
                        escribirLog("fldBatch: " + ex.Message);
                    }
                   
                }
                    //Proceso en la pestaña del batch de autorretenciones, para traer los check
                    if (pVal.FormTypeEx == "SCL_ParamIniLoc" && (pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "fldBatch") && pVal.Before_Action == true)
                {
                    try
                    {
                        string Agrup = string.Empty;
                        string AsRes = string.Empty;
                        string AutB = string.Empty;
                        string query = string.Empty;
                       
                        query = "SELECT \"U_SCL_Agrup\",\"U_SCL_AsRes\",\"U_SCL_AutB\" FROM \"@SCL_LOC_VERSION\"";
                        SAPbobsCOM.Recordset oRecPro = null;
                        oRecPro = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                        oRecPro.DoQuery(query);
                        if (oRecPro.RecordCount > 0) {
                            Agrup = oRecPro.Fields.Item(0).Value;
                            AsRes= oRecPro.Fields.Item(1).Value;
                            AutB = oRecPro.Fields.Item(2).Value;


                            try
                            {
                                if (AutB.Equals("Y"))
                                {
                                    oNewItem = oForm.Items.Add("chkAutB", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                                    oNewItem.ToPane = 3;
                                    oNewItem.FromPane = 3;
                                    oNewItem.Left = 103;
                                    oNewItem.Top = 76;
                                    oNewItem.Height = 14;
                                    oNewItem.Width = 123;
                                    oNewItem.Visible = true;
                                    oChekBox = ((CheckBox)(oNewItem.Specific));
                                    oChekBox.Caption = "Autorretención en batch";
                                    oChekBox.ValOn = "Y";
                                    oChekBox.ValOff = "N";
                                    oChekBox.DataBind.SetBound(true, "@SCL_LOC_VERSION", "U_SCL_AutB");
                                    oChekBox.Checked = true;
                                }
                                else
                                {
                                    oNewItem = oForm.Items.Add("chkAutB", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                                    oNewItem.ToPane = 3;
                                    oNewItem.FromPane = 3;
                                    oNewItem.Left = 103;
                                    oNewItem.Top = 76;
                                    oNewItem.Height = 14;
                                    oNewItem.Width = 123;
                                    oNewItem.Visible = true;
                                    oChekBox = ((CheckBox)(oNewItem.Specific));
                                    oChekBox.Caption = "Autorretención en batch";
                                    oChekBox.ValOn = "Y";
                                    oChekBox.ValOff = "N";
                                    oChekBox.DataBind.SetBound(true, "@SCL_LOC_VERSION", "U_SCL_AutB");
                                    oChekBox.Checked = false;
                                }


                                if (AsRes.Equals("Y"))
                                {

                                    oNewItem = oForm.Items.Add("chkAsRes", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                                    oNewItem.ToPane = 3;
                                    oNewItem.FromPane = 3;
                                    oNewItem.Left = 103;
                                    oNewItem.Top = 101;
                                    oNewItem.Height = 14;
                                    oNewItem.Width = 123;
                                    oNewItem.Visible = true;
                                    oChekBox = ((CheckBox)(oNewItem.Specific));
                                    oChekBox.Caption = "Asiento resumido";
                                    oChekBox.ValOn = "Y";
                                    oChekBox.ValOff = "N";
                                    oChekBox.DataBind.SetBound(true, "@SCL_LOC_VERSION", "U_SCL_AsRes");
                                    oChekBox.Checked = true;
                                }
                                else
                                {
                                    oNewItem = oForm.Items.Add("chkAsRes", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                                    oNewItem.ToPane = 3;
                                    oNewItem.FromPane = 3;
                                    oNewItem.Left = 103;
                                    oNewItem.Top = 101;
                                    oNewItem.Height = 14;
                                    oNewItem.Width = 123;
                                    oNewItem.Visible = true;
                                    oChekBox = ((CheckBox)(oNewItem.Specific));
                                    oChekBox.Caption = "Asiento resumido";
                                    oChekBox.ValOn = "Y";
                                    oChekBox.ValOff = "N";
                                    oChekBox.DataBind.SetBound(true, "@SCL_LOC_VERSION", "U_SCL_AsRes");
                                    oChekBox.Checked = false;
                                }
                                if (Agrup.Equals("Y"))
                                {
                                    oNewItem = oForm.Items.Add("chkAgrup", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                                    oNewItem.ToPane = 3;
                                    oNewItem.FromPane = 3;
                                    oNewItem.Left = 103;
                                    oNewItem.Top = 126;
                                    oNewItem.Height = 14;
                                    oNewItem.Width = 123;
                                    oNewItem.Visible = true;
                                    oChekBox = ((CheckBox)(oNewItem.Specific));
                                    oChekBox.Caption = "Agrupación SN";
                                    oChekBox.ValOn = "Y";
                                    oChekBox.ValOff = "N";
                                    oChekBox.DataBind.SetBound(true, "@SCL_LOC_VERSION", "U_SCL_Agrup");
                                    oChekBox.Checked = true;
                                }
                                else
                                {
                                    oNewItem = oForm.Items.Add("chkAgrup", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                                    oNewItem.ToPane = 3;
                                    oNewItem.FromPane = 3;
                                    oNewItem.Left = 103;
                                    oNewItem.Top = 126;
                                    oNewItem.Height = 14;
                                    oNewItem.Width = 123;
                                    oNewItem.Visible = true;
                                    oChekBox = ((CheckBox)(oNewItem.Specific));
                                    oChekBox.Caption = "Agrupación SN";
                                    oChekBox.ValOn = "Y";
                                    oChekBox.ValOff = "N";
                                    oChekBox.DataBind.SetBound(true, "@SCL_LOC_VERSION", "U_SCL_Agrup");
                                    oChekBox.Checked = false;
                                }

                                oForm.Refresh();
                            }
                            catch (Exception ex)
                            {

                            }

                            
                            
                        }

                    }
                    catch (Exception ex)
                    {
                        SBO_Application.MessageBox(ex.Message);
                        escribirLog("fldBatch: " + ex.Message);
                    }
                }																			  
            }
            catch (Exception ex)
            {
                Business.escribirLog("EventFieldParamLoc: " + ex.Message);
                //SBO_Application.MessageBox(ex.Message);
            }
        }

        public static void cargarGrid(string formulario, string FECHAINI, string FECHAFIN)
        {
            try
            {
                SAPbouiCOM.GridColumn oGridColumn;
                string sSQL = string.Format(Consultas.Default.Cierre, FECHAINI, FECHAFIN);
                oForm = SBO_Application.Forms.Item(formulario);

                oItem = oForm.Items.Item("Grid");
                oGrid = oItem.Specific;
                oForm.DataSources.DataTables.Item(0).ExecuteQuery(sSQL);
                oGrid.DataTable = oForm.DataSources.DataTables.Item("DTCIERRE");
                //oGrid.Item.Enabled = false;

                for (int i = 0; i < oGrid.Columns.Count; i++)
                {
                    if (i == oGrid.Columns.Count - 1)
                    {
                        oGridColumn = oGrid.Columns.Item(i);
                        oGridColumn.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                        oGridColumn.Editable = true;
                    }
                    else
                    {
                        oGridColumn = oGrid.Columns.Item(i);
                        oGridColumn.Editable = false;
                    }
                }
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
                escribirLog("CargarGrid: " + ex.Message);
            }
        }

        /// <summary>
        /// Cargue de datos DatosGlobalesFE
        /// </summary>
        public static void cargarDatosGlobalesSAP()
        {
            try
            {
                SAPbobsCOM.IAdminInfo oAdminInfo;

                oCmpSrv = oCompany.GetCompanyService();
                oAdminInfo = oCmpSrv.GetAdminInfo();
                DatosGlobalesFE.sapdecimal = oAdminInfo.DecimalSeparator;
                DatosGlobalesFE.sapMillar = oAdminInfo.ThousandsSeparator;

                DatosGlobalesFE.localdecimal = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;
                DatosGlobalesFE.localMillar = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberGroupSeparator;

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oAdminInfo);
                oAdminInfo = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCmpSrv);
                oCmpSrv = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
                Business.escribirLog("CargarDatosGlobalesSAP: " + ex.Message);
            }
        }

        /// <summary>
        /// Definicion Timers
        /// Verificar Estados
        /// Reenvio Documentos
        /// </summary>
        public static void startMonitorSAPB1()
        {

            #region TimerReSend
            // Alternate method: create a Timer with an interval argument to the constructor.
            //aTimer = new System.Timers.Timer(2000);

            // Create a timer with a five second interval.

            if (DatosGlobalesFE.TimerReenvio != 0)
            {
                aTimer = new System.Timers.Timer(DatosGlobalesFE.TimerReenvio * 60000);
                aTimer.Enabled = true;
            }
            else
            {
                aTimer = new System.Timers.Timer(10 * 60000);
                aTimer.Enabled = false;
            }
            // Hook up the Elapsed event for the timer. 
            aTimer.Elapsed += procesoReEnvio;

            // Have the timer fire repeated events (true is the default)
            aTimer.AutoReset = true;

            // Start the timer
            //aTimer.Enabled = true;
            #endregion TimerReSend


            #region TimerVerificaEstados
            // Alternate method: create a Timer with an interval argument to the constructor.
            //aTimer = new System.Timers.Timer(2000);

            // Create a timer with a five second interval.
            if (DatosGlobalesFE.TimerEstado != 0)
            {
                bTimer = new System.Timers.Timer(DatosGlobalesFE.TimerEstado * 60000);
                bTimer.Enabled = true;
            }
            else
            {
                bTimer = new System.Timers.Timer(10 * 60000);
                bTimer.Enabled = false;
            }

            // Hook up the Elapsed event for the timer. 
            bTimer.Elapsed += procesoVerificarEstado;

            // Have the timer fire repeated events (true is the default)
            bTimer.AutoReset = true;

            // Start the timer
            //bTimer.Enabled = true;
            #endregion TimerVerificaEstados
        }

        /// <summary>
        /// Proceso Timer Verificar Estado
        /// </summary>
        public static void procesoReEnvio(Object source, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                if (flagReSend == true)
                {
                    flagReSend = false;
                    autoReEnvio();
                    flagReSend = true;
                }
            }
            catch (Exception ex)
            {
                escribirLog("TimerReEnVio: " + ex.Message);
                flagReSend = true;
            }
        }

        /// <summary>
        /// Proceso Timer ReEnvio
        /// </summary>
        public static void procesoVerificarEstado(Object source, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                if (flagVerifiStatus == true)
                {
                    flagVerifiStatus = false;
                    verificarEstados();
                    flagVerifiStatus = true;
                }
            }
            catch (Exception ex)
            {
                escribirLog("TimerVerificarEstados: " + ex.Message);
                flagVerifiStatus = true;
            }
        }

        /// <summary>
        /// Envio de documentos al web service de facturacion electronica
        /// </summary>
        /// <param name="query">Consulta para extraer los datos del documento</param>
        /// <param name="docEntry">Identificador del documento</param>
        /// <param name="obj">Tipo de objeto del documento</param>
        public static void envioDocumento(string query, int docEntry, string obj, double valorTotal)
        {
        //    //try
        //    ////{
        //    //    string valorEnLetras = "";
        //    //    string consulta = "";
        //    //    valorEnLetras = NumerosALetras.enletras(valorTotal.ToString());
        //    //    consulta = string.Format(query, docEntry, valorEnLetras);
        //    //    string txtfac = generarTXT(consulta);
        //    //    escribirLog("TxtGenerado: " + txtfac);
        //    //    var xml = WebServiceController.procesartxt(txtfac);

        //    //    var doc = new XmlDocument();
        //    //    doc.LoadXml(xml);

        //    //    var docresult = new XmlDocument();
        //    //    docresult.LoadXml(doc.GetElementsByTagName("procesarTextoPlanoResult")[0].InnerText);

        //    //    bool agrego = actualizarCamposFE(obj, docEntry, docresult);
        //    //}
        //    //catch (Exception ex)
        //    //{
        //    //    escribirLog("envioDocumento: " + ex.Message);
        //    //}
        }

        /// <summary>
        /// Envio de documentos al web service de facturacion electronica
        /// </summary>
        /// <param name="query">Consulta para extraer los datos del documento</param>
        /// <param name="docEntry">Identificador del documento</param>
        /// <param name="obj">Tipo de objeto del documento</param>
        public static void verificarDIAN(string cufe, int docEntry, string obj)
        {
            try
            {
     
            }
            catch (Exception ex)
            {
                escribirLog("verificarDIAN: " + ex.Message);
            }
        }

        /// <summary>
        /// Obtiene la consulta para extraer los datos segun el documento
        /// </summary>
        /// <param name="seriesName">Nombre del la serie de nuemracion para validar prefijo</param>
        public static string obtenerConsulta(string seriesName, string oForm)
        {
            try
            {
                string nombreserie = "";
                nombreserie = seriesName;
                string[] ArrLine;
                string delimStr = "_";
                char[] delimiter = delimStr.ToCharArray();
                int x = 2;
                ArrLine = nombreserie.Split(delimiter, x);

                if (oForm == "13")
                {
                    if (ArrLine[0] == DatosGlobalesFE.IdentificadorFE)
                    {
                        return Consultas.Default.FacturaVenta;
                    }
                    else if (ArrLine[0] == DatosGlobalesFE.IdentificadorFEC)
                    {
                        return Consultas.Default.FacturaConti;
                    }
                    else if (ArrLine[0] == DatosGlobalesFE.IdentificadorFEX)
                    {
                        return Consultas.Default.FacturaExpo;
                    }
                    else
                    {
                        return null;
                    }
                }
                else if (oForm == "65303")
                {
                    return Consultas.Default.NotaDebito;
                }
                else if (oForm == "14")
                {
                    return Consultas.Default.NotaCredito;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                escribirLog("obtenerConsulta: " + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Genera el arcvhio txt con la estructura para el envio
        /// </summary>
        /// <param name="query">Consulta con la cula sera extraida la informacion del docuemnto a emitir</param>
        public static string generarTXT(string query)
        {
            try
            {
                string myStr = "";
                System.Data.DataTable sendFile = new System.Data.DataTable();
                int i = 0;

                Recordset oRecordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string sSql = query;
                oRecordset.DoQuery(sSql);

                sendFile = convertRS2DT(oRecordset);

                using (MemoryStream ms = new MemoryStream())
                {
                    StreamWriter sw = new StreamWriter(ms);
                    foreach (DataRow row in sendFile.Rows)
                    {
                        object[] array = row.ItemArray;

                        for (i = 0; i < array.Length - 1; i++)
                        {
                            sw.Write(array[i].ToString());
                        }
                        sw.WriteLine(array[i].ToString());
                        //sw.WriteLine();
                    }
                    sw.Flush();
                    ms.Position = 0;
                    StreamReader sr = new StreamReader(ms);
                    myStr = sr.ReadToEnd();
                }
                string text = myStr;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
                oRecordset = null;
                GC.Collect();
                return text;
            }
            catch (Exception ex)
            {
                escribirLog("Strtxt: " + ex.Message);
                return "";
            }
        }

        /// <summary>
        /// Actualzia campos de usaurio documento emitido
        /// </summary>
        /// <param name="objType">Tipo de objeto</param>
        /// <param name="docEntry">Numero interno del docuemtno</param>
        /// <param name="valor">xml respuesta del operador tecnologico</param>
        public static bool actualizarCamposFE(string objType, int docEntry, XmlDocument valor)
        {
            try
            {
                if (objType == "13")
                {
                    Documents oInvoice = oCompany.GetBusinessObject(BoObjectTypes.oInvoices);
                    if (oInvoice.GetByKey(docEntry))
                    {
                        string docNum = "";
                        docNum = Convert.ToString(oInvoice.DocNum);
                        escribirLog("XML_Respuesta: " + valor.InnerXml);
                        string documento = "";
                        BoDocumentSubType SubType;
                        BoYesNoEnum FacReserva;
                        SubType = oInvoice.DocumentSubType;
                        FacReserva = oInvoice.ReserveInvoice;

                        if (SubType == BoDocumentSubType.bod_ExportInvoice)
                        {
                            documento = "Factura de Exportacion";
                        }
                        else if (SubType == BoDocumentSubType.bod_DebitMemo)
                        {
                            documento = "Nota de Debito";
                        }
                        else if (SubType == BoDocumentSubType.bod_None && FacReserva == BoYesNoEnum.tYES)
                        {
                            documento = "Factura de Reserva";
                        }
                        else
                        {
                            documento = "Factura de Venta";
                        }

                        if (valor.GetElementsByTagName("ErrorCode")[0] != null && !string.IsNullOrEmpty(valor.GetElementsByTagName("ErrorCode")[0].InnerText))
                        {
                            escribirLog("ErrorCode: " + valor.GetElementsByTagName("ErrorCode")[0].InnerText);
                            oInvoice.UserFields.Fields.Item("U_SCL_FE_Estado").Value = "1";
                        }

                        if (valor.GetElementsByTagName("cMensaje")[0] != null && !string.IsNullOrEmpty(valor.GetElementsByTagName("cMensaje")[0].InnerText))
                        {
                            string estado = "";
                            string mensaje = "";
                            string acuseDIAN = "";
                            estado = valor.GetElementsByTagName("cEstatus")[0].InnerText;
                            mensaje = valor.GetElementsByTagName("cMensaje")[0].InnerText;

                            if (valor.GetElementsByTagName("cXMLAcuse")[0] != null && !string.IsNullOrEmpty(valor.GetElementsByTagName("cXMLAcuse")[0].InnerText))
                            {
                                acuseDIAN = valor.GetElementsByTagName("cXMLAcuse")[0].InnerText;
                                mensaje = mensaje + ".\n" + Decode64tostring(acuseDIAN);
                            }
                            if (valor.GetElementsByTagName("cMensaje")[0].InnerText == "EXITOSA")
                            {
                                oInvoice.UserFields.Fields.Item("U_SCL_FE_Estado").Value = "3";
                                estado = "EXITOSA";
                            }
                            else if (estado == "102")
                            {
                                oInvoice.UserFields.Fields.Item("U_SCL_FE_Estado").Value = "2";
                                oInvoice.UserFields.Fields.Item("U_SCL_FE_RESULT").Value = mensaje;
                            }
                            else
                            {
                                oInvoice.UserFields.Fields.Item("U_SCL_FE_Estado").Value = "4";
                                oInvoice.UserFields.Fields.Item("U_SCL_FE_RESULT").Value = mensaje;
                            }

                            sendMessage(objType, documento, docNum, docEntry.ToString(), estado, mensaje, true);
                        }

                        if (valor.GetElementsByTagName("XMLTimbrado")[0] != null && !string.IsNullOrEmpty(valor.GetElementsByTagName("XMLTimbrado")[0].InnerText))
                        {
                            escribirLog("XMLTimbrado: " + valor.GetElementsByTagName("XMLTimbrado")[0].InnerText);
                            string xmlTimbrado = "";
                            xmlTimbrado = valor.GetElementsByTagName("XMLTimbrado")[0].InnerXml;
                            string CUFE = "";
                            CUFE = valor.GetElementsByTagName("cbc:UUID")[0].InnerXml;
                            escribirLog("CUFE: " + valor.GetElementsByTagName("cbc:UUID")[0].InnerXml);
                            oInvoice.UserFields.Fields.Item("U_SCL_FE_Estado").Value = "2";
                            oInvoice.UserFields.Fields.Item("U_SCL_FE_RESULT").Value = xmlTimbrado;
                            oInvoice.UserFields.Fields.Item("U_SCL_FE_CUFE").Value = CUFE;
                            sendMessage(objType, documento, docNum, docEntry.ToString(), "OK", CUFE, false);
                        }

                        if (valor.GetElementsByTagName("ErrorMessage")[0] != null && !string.IsNullOrEmpty(valor.GetElementsByTagName("ErrorMessage")[0].InnerText))
                        {
                            escribirLog("ErrorMessage: " + valor.GetElementsByTagName("ErrorMessage")[0].InnerText);
                            int x = 2;
                            string mensajeError = "";
                            mensajeError = valor.GetElementsByTagName("ErrorMessage")[0].InnerText;
                            string[] ArrLine;
                            string delimStr = "\n";
                            char[] delimiter = delimStr.ToCharArray();
                            ArrLine = mensajeError.Split(delimiter, x);
                            mensajeError = ArrLine[0];

                            oInvoice.UserFields.Fields.Item("U_SCL_FE_RESULT").Value = mensajeError;

                            sendMessage(objType, documento, docNum, docEntry.ToString(), "Error", mensajeError, false);
                        }

                        if (oInvoice.Update() != 0)
                            escribirLog("Documento:" + documento + " DocNum: " + docNum + " Error: " + oCompany.GetLastErrorDescription());
                        //SBO_Application.MessageBox(oCompany.GetLastErrorDescription());
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvoice);
                    oInvoice = null;
                    GC.Collect();
                }
                else if (objType == "14")
                {
                    Documents oCreditNote = oCompany.GetBusinessObject(BoObjectTypes.oCreditNotes);
                    if (oCreditNote.GetByKey(docEntry))
                    {
                        string docNum = "";
                        docNum = Convert.ToString(oCreditNote.DocNum);
                        escribirLog("XML_Respuesta: " + valor.InnerXml);
                        string documento = "";
                        BoDocumentSubType SubType;
                        BoYesNoEnum FacReserva;
                        SubType = oCreditNote.DocumentSubType;
                        FacReserva = oCreditNote.ReserveInvoice;

                        if (SubType == BoDocumentSubType.bod_ExportInvoice)
                        {
                            documento = "Factura de Exportacion";
                        }
                        else if (SubType == BoDocumentSubType.bod_DebitMemo)
                        {
                            documento = "Nota de Debito";
                        }
                        else if (SubType == BoDocumentSubType.bod_None && FacReserva == BoYesNoEnum.tYES)
                        {
                            documento = "Factura de Reserva";
                        }
                        else
                        {
                            documento = "Nota de Credito";
                        }

                        if (valor.GetElementsByTagName("ErrorCode")[0] != null && !string.IsNullOrEmpty(valor.GetElementsByTagName("ErrorCode")[0].InnerText))
                        {
                            escribirLog("ErrorCode: " + valor.GetElementsByTagName("ErrorCode")[0].InnerText);
                            oCreditNote.UserFields.Fields.Item("U_SCL_FE_Estado").Value = "1";
                        }

                        if (valor.GetElementsByTagName("cMensaje")[0] != null && !string.IsNullOrEmpty(valor.GetElementsByTagName("cMensaje")[0].InnerText))
                        {
                            string estado = "";
                            string mensaje = "";
                            string acuseDIAN = "";
                            estado = valor.GetElementsByTagName("cEstatus")[0].InnerText;
                            mensaje = valor.GetElementsByTagName("cMensaje")[0].InnerText;

                            if (valor.GetElementsByTagName("cXMLAcuse")[0] != null && !string.IsNullOrEmpty(valor.GetElementsByTagName("cXMLAcuse")[0].InnerText))
                            {
                                acuseDIAN = valor.GetElementsByTagName("cXMLAcuse")[0].InnerText;
                                mensaje = mensaje + ".\n" + Decode64tostring(acuseDIAN);
                            }
                            if (valor.GetElementsByTagName("cMensaje")[0].InnerText == "EXITOSA")
                            {
                                oCreditNote.UserFields.Fields.Item("U_SCL_FE_Estado").Value = "3";
                                estado = "EXITOSA";
                            }
                            else if (estado == "102")
                            {
                                oCreditNote.UserFields.Fields.Item("U_SCL_FE_Estado").Value = "2";
                                oCreditNote.UserFields.Fields.Item("U_SCL_FE_RESULT").Value = mensaje;
                            }
                            else
                            {
                                oCreditNote.UserFields.Fields.Item("U_SCL_FE_Estado").Value = "4";
                                oCreditNote.UserFields.Fields.Item("U_SCL_FE_RESULT").Value = mensaje;
                            }

                            sendMessage(objType, documento, docNum, docEntry.ToString(), estado, mensaje, true);
                        }

                        if (valor.GetElementsByTagName("XMLTimbrado")[0] != null && !string.IsNullOrEmpty(valor.GetElementsByTagName("XMLTimbrado")[0].InnerText))
                        {
                            escribirLog("XMLTimbrado: " + valor.GetElementsByTagName("XMLTimbrado")[0].InnerText);
                            string xmlTimbrado = "";
                            xmlTimbrado = valor.GetElementsByTagName("XMLTimbrado")[0].InnerXml;
                            string CUFE = "";
                            CUFE = valor.GetElementsByTagName("cbc:UUID")[0].InnerXml;
                            escribirLog("CUFE: " + valor.GetElementsByTagName("cbc:UUID")[0].InnerXml);
                            oCreditNote.UserFields.Fields.Item("U_SCL_FE_Estado").Value = "2";
                            oCreditNote.UserFields.Fields.Item("U_SCL_FE_RESULT").Value = xmlTimbrado;
                            oCreditNote.UserFields.Fields.Item("U_SCL_FE_CUFE").Value = CUFE;
                            sendMessage(objType, documento, docNum, docEntry.ToString(), "OK", CUFE, false);
                        }
                        if (valor.GetElementsByTagName("ErrorMessage")[0] != null && !string.IsNullOrEmpty(valor.GetElementsByTagName("ErrorMessage")[0].InnerText))
                        {
                            escribirLog("ErrorMessage: " + valor.GetElementsByTagName("ErrorMessage")[0].InnerText);
                            int x = 2;
                            string mensajeError = "";
                            mensajeError = valor.GetElementsByTagName("ErrorMessage")[0].InnerText;
                            string[] ArrLine;
                            string delimStr = "\n";
                            char[] delimiter = delimStr.ToCharArray();
                            ArrLine = mensajeError.Split(delimiter, x);
                            mensajeError = ArrLine[0];

                            oCreditNote.UserFields.Fields.Item("U_SCL_FE_RESULT").Value = mensajeError;

                            sendMessage(objType, documento, docNum, docEntry.ToString(), "Error", mensajeError, false);
                        }
                        if (oCreditNote.Update() != 0)
                            escribirLog("Documento:" + documento + " DocNum: " + docNum + " Error: " + oCompany.GetLastErrorDescription());
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCreditNote);
                    oCreditNote = null;
                    GC.Collect();
                }
                return false;
            }
            catch (Exception ex)
            {
                escribirLog("actualizarCampos: " + ex.Message);
                if (oCompany.InTransaction)
                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                return false;
            }
        }

        /// <summary>
        /// Enviar mensaje SAP
        /// </summary>
        /// <param name="objType">Tipo de objeto</param>
        /// <param name="docEntry">Numero interno del docuemtno</param>
        /// <param name="valor">xml respuesta del operador tecnologico</param>
        public static void sendMessage(string objType, string documento, string docNum, string docEntry, string status, string msg, bool verificarEstado)
        {
            try
            {
                SAPbobsCOM.Message oMessage = null;
                MessagesService oMessageService = null;
                MessageDataColumns pMessageDataColumns = null;
                MessageDataColumn pMessageDataColumn = null;
                MessageDataLines oLines = null;
                MessageDataLine oLine = null;
                RecipientCollection oRecipientCollection = null;

                //get company service
                oCmpSrv = oCompany.GetCompanyService();

                //get msg service
                oMessageService = (SAPbobsCOM.MessagesService)oCmpSrv.GetBusinessService(ServiceTypes.MessagesService);

                // get the data interface for the new message
                oMessage = ((SAPbobsCOM.Message)(oMessageService.GetDataInterface(MessagesServiceDataInterfaces.msdiMessage)));

                // fill subject

                if (!verificarEstado)
                {
                    if (status == "OK")
                    {
                        oMessage.Subject = "Envio satisfactorio " + documento + " " + docNum;
                        oMessage.Priority = BoMsgPriorities.pr_Normal;
                        oMessage.Text = msg;
                    }
                    else
                    {
                        oMessage.Subject = "Error Envio " + documento + " " + docNum;
                        oMessage.Priority = BoMsgPriorities.pr_High;
                        oMessage.Text = msg;
                    }
                }
                else
                {
                    if (status == "EXITOSA")
                    {
                        oMessage.Subject = "Verificacion DIAN: " + status + " " + documento + " " + docNum;
                        oMessage.Priority = BoMsgPriorities.pr_Normal;
                        oMessage.Text = msg;
                    }
                    else
                    {
                        oMessage.Subject = "Verificacion DIAN: " + "Fallida" + " " + documento + " " + docNum;
                        oMessage.Priority = BoMsgPriorities.pr_High;
                        oMessage.Text = msg;
                    }
                }

                // Add Recipient 
                oRecipientCollection = oMessage.RecipientCollection;

                oRecipientCollection.Add();

                // send internal message
                oRecipientCollection.Item(0).SendInternal = BoYesNoEnum.tYES;

                // add existing user name
                oRecipientCollection.Item(0).UserCode = SBO_Application.Company.UserName;

                // get columns data
                pMessageDataColumns = oMessage.MessageDataColumns;
                // get column
                pMessageDataColumn = pMessageDataColumns.Add();
                // set column name
                pMessageDataColumn.ColumnName = "Documento";
                // set link to a real object in the application
                pMessageDataColumn.Link = BoYesNoEnum.tNO;
                // get lines
                oLines = pMessageDataColumn.MessageDataLines;
                // add new line
                oLine = oLines.Add();
                // set the line value
                oLine.Value = documento;


                pMessageDataColumn = pMessageDataColumns.Add();
                // set column name
                pMessageDataColumn.ColumnName = "Numero Documento";
                // set link to a real object in the application
                pMessageDataColumn.Link = BoYesNoEnum.tYES;
                // get lines
                oLines = pMessageDataColumn.MessageDataLines;
                // add new line
                oLine = oLines.Add();
                // set the line value
                oLine.Value = docNum;
                // set the link to BusinessPartner (the object type for Bp is 2)
                oLine.Object = objType;
                // set the Bp code
                oLine.ObjectKey = docEntry;
                oLine.Value = docNum;

                pMessageDataColumn = pMessageDataColumns.Add();
                // set column name
                pMessageDataColumn.ColumnName = "Estado";
                // set link to a real object in the application
                pMessageDataColumn.Link = BoYesNoEnum.tNO;
                // get lines
                oLines = pMessageDataColumn.MessageDataLines;
                // add new line
                oLine = oLines.Add();
                // set the line value
                oLine.Value = status;


                if (status == "OK")
                {
                    pMessageDataColumn = pMessageDataColumns.Add();
                    // set column name
                    pMessageDataColumn.ColumnName = "CUFE";
                    // set link to a real object in the application
                    pMessageDataColumn.Link = BoYesNoEnum.tNO;
                    // get lines
                    oLines = pMessageDataColumn.MessageDataLines;
                    // add new line
                    oLine = oLines.Add();
                    // set the line value
                    oLine.Value = msg;
                }
                else
                {
                    pMessageDataColumn = pMessageDataColumns.Add();
                    // set column name
                    pMessageDataColumn.ColumnName = "Mensaje";
                    // set link to a real object in the application
                    pMessageDataColumn.Link = BoYesNoEnum.tNO;
                    // get lines
                    oLines = pMessageDataColumn.MessageDataLines;
                    // add new line
                    oLine = oLines.Add();
                    // set the line value
                    oLine.Value = msg;
                }

                // send the message
                oMessageService.SendMessage(oMessage);
            }
            catch (Exception ex)
            {
                escribirLog("EnviarMensaje: " + ex.Message);
                if (Business.oCompany.InTransaction)
                    Business.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                //return false;
            }
        }

        /// <summary>
        /// Registrar log en archvio txt
        /// </summary>
        /// <param name="cadenalog">mensaje a escribir en el archvio de log</param>
        public static void escribirLog(string cadenalog)
        {
            try
            {
                oConnection.crearCarpeta();
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
                ////System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + (fullPath) + "\\" + ArchivoLog;

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

        /// <summary>
        /// Metodo para verificar el estado de los documentos emitidos
        /// </summary>
        public static void verificarEstados()
        {
            try
            {
                System.Data.DataTable ResultQuery = new System.Data.DataTable();
                Recordset oRecordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string sSql = string.Format(Consultas.Default.Estados);
                oRecordset.DoQuery(sSql);

                if (oRecordset.RecordCount > 0)
                {
                    ResultQuery = convertRS2DT(oRecordset);

                    for (int i = 0; i < ResultQuery.Rows.Count; i++) //Looping through rows
                    {
                        string cufe = "";
                        string objType = "";
                        int docentry;

                        cufe = Convert.ToString(ResultQuery.Rows[i]["CUFE"]);
                        objType = Convert.ToString(ResultQuery.Rows[i]["ObjType"]);
                        docentry = Convert.ToInt32(ResultQuery.Rows[i]["DocEntry"]);

                        verificarDIAN(cufe, docentry, objType);
                    }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
                oRecordset = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                escribirLog("verifystatus: " + ex.Message);
            }
        }

        /// <summary>
        /// Metodo para reenviar los documentos emitidos que quedan en error
        /// </summary>
        public static void autoReEnvio()
        {
            try
            {
                System.Data.DataTable ResultQuery = new System.Data.DataTable();
                Recordset oRecordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string sSql = string.Format(Consultas.Default.ReEnvio);
                oRecordset.DoQuery(sSql);

                if (oRecordset.RecordCount > 0)
                {
                    ResultQuery = convertRS2DT(oRecordset);

                    for (int i = 0; i < ResultQuery.Rows.Count; i++) //Looping through rows
                    {
                        string consulta = "";
                        int docEntry;
                        string objType = "";
                        double docTotal = 0;

                        consulta = obtenerConsulta(Convert.ToString(ResultQuery.Rows[i]["SeriesName"]), Convert.ToString(ResultQuery.Rows[i]["TipoDoc"]));
                        docEntry = Convert.ToInt32(ResultQuery.Rows[i]["DocEntry"]);
                        objType = Convert.ToString(ResultQuery.Rows[i]["ObjType"]);
                        docTotal = Convert.ToDouble(ResultQuery.Rows[i]["DocTotal"]);

                        envioDocumento(consulta, docEntry, objType, docTotal);
                    }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
                oRecordset = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                escribirLog("AutoReSend: " + ex.Message);
            }
        }

        /// <summary>
        /// Metodo para almacenar recorset en un datatable
        /// </summary>
        /// <param name="RS">recorset que se va a almacenar en DataTable</param>
        public static System.Data.DataTable convertRS2DT(Recordset RS)
        {
            try
            {
                System.Data.DataTable dtTable = new System.Data.DataTable();
                System.Data.DataColumn NewCol = default(System.Data.DataColumn);
                DataRow NewRow = default(DataRow);
                int ColCount = 0;

                //try
                //{

                while (ColCount < RS.Fields.Count)
                {
                    string dataType = "System.";
                    switch (RS.Fields.Item(ColCount).Type)
                    {
                        case BoFieldTypes.db_Alpha:
                            dataType = dataType + "String";
                            break;
                        case BoFieldTypes.db_Date:
                            dataType = dataType + "DateTime";
                            break;
                        case BoFieldTypes.db_Float:
                            dataType = dataType + "Double";
                            break;
                        case BoFieldTypes.db_Memo:
                            dataType = dataType + "String";
                            break;
                        case BoFieldTypes.db_Numeric:
                            dataType = dataType + "Decimal";
                            break;
                        default:
                            dataType = dataType + "String";
                            break;
                    }

                    NewCol = new System.Data.DataColumn(RS.Fields.Item(ColCount).Name, System.Type.GetType(dataType));
                    dtTable.Columns.Add(NewCol);
                    ColCount++;
                }
                int iCol = 0;
                while (!(RS.EoF))
                {
                    NewRow = dtTable.NewRow();

                    dtTable.Rows.Add(NewRow);

                    iCol = 0;
                    ColCount = 0;
                    while (ColCount < RS.Fields.Count)
                    {
                        //NewRow.Item(RS.Fields.Item(ColCount).Name) = RS.Fields.Item(ColCount).Value;
                        NewRow[iCol] = RS.Fields.Item(ColCount).Value;
                        iCol++;
                        ColCount++;
                    }
                    RS.MoveNext();
                }
                return dtTable;
            }
            catch (Exception ex)
            {
                escribirLog("RSToDataTable: " + ex.Message);
                //SBO_Application.MessageBox(ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Metodo para decodificar base64
        /// </summary>
        /// <param name="toDecode">base64 a decodificar</param>
        public static string Decode64tostring(string toDecode)
        {
            try
            {
                byte[] data = Convert.FromBase64String(toDecode);
                string decodedString = Encoding.UTF8.GetString(data);
                return decodedString;
            }
            catch (Exception ex)
            {
                escribirLog("Decode64tostring: " + ex.Message);
                //SBO_Application.MessageBox(ex.Message);
                return null;
            }
        }


        //public string ConexionServiceLayer()
        //{
        //    oConnection.CredencialesSL();
        //    //ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);
        //    var cliente = new RestClient(ConfigurationManager.AppSettings["SLAddress"].ToString());
        //    //string CompanyDB = ConfigurationManager.AppSettings["CompanyDB"].ToString();
        //    string CompanyDB = oCompany.CompanyDB;
        //    string Password = ConfigurationManager.AppSettings["Password"].ToString();
        //    string UserName = ConfigurationManager.AppSettings["UserName"].ToString();
        //    //string sessionID = string.Empty;
        //    var data = new Dictionary<string, string>
        //    {
        //        {"CompanyDB", (CompanyDB) },
        //        { "Password", (Password) },
        //        {"UserName",  (UserName) }
        //    };
        //    var body = JsonConvert.SerializeObject(data);
        //    var request = new RestRequest("/b1s/v1/Login", Method.POST);
        //    request.RequestFormat = DataFormat.Json;
        //    request.AddParameter("application/json", body, ParameterType.RequestBody);
        //    RestResponse response = (RestResponse)cliente.Execute(request);
        //    int status = (int)response.StatusCode;
        //    if (response.StatusCode.Equals(HttpStatusCode.OK))
        //    {
        //        dynamic dyn = JsonConvert.DeserializeObject(response.Content);
        //        foreach (var obj in dyn)
        //        {
        //            if (obj.Name.Equals("SessionId"))
        //                sessionID = obj.Value;
        //        }
        //    }
        //    else if (response.StatusCode.Equals(HttpStatusCode.NotFound))
        //    {
        //        SBO_Application.MessageBox("Error con la URL de Service Layer");
        //    }
        //    return sessionID;
        //}


        public void AsignarTerceroAsientos()
        {
            SAPbobsCOM.Recordset oRctAsientos;
            SAPbobsCOM.Recordset oRctNumLineas;
            SAPbouiCOM.ProgressBar barraProgreso;
            try
            {
                oRctAsientos = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = Properties.Resources.IdAsientoDoc;
                oRctAsientos.DoQuery(query);                
                barraProgreso = SBO_Application.StatusBar.CreateProgressBar("Barra de progreso", oRctAsientos.RecordCount, false);
                //oConnection.ConCompany(oCompany, SBO_Application);
                string sessionID = oConnection.ConexionServiceLayer();

                //AÑADIR TRANSACCION ID, PARA REALIZAR LA ACTUALIZACION DEL CAMPO EMPLEANDO EL SERVICE LAYER...
                while (!oRctAsientos.EoF)
                {
                    List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
                    Dictionary<string, object> row;
                    //string lineaTercero = oRctAsientos.Fields.Item("U_SCL_CodeSN").Value.ToString();
                    string TransId = oRctAsientos.Fields.Item("TransId").Value.ToString();
                    string cardCode = oRctAsientos.Fields.Item("ShortName").Value.ToString();
                    oRctNumLineas = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string queryLineas = Properties.Resources.NumLineasAsiento;                    
                    oRctNumLineas.DoQuery(queryLineas.Replace("%", TransId));
                    int numLineas = Convert.ToInt32(oRctNumLineas.Fields.Item("NUMFILAS").Value);
                    foreach (var index in Enumerable.Range(1, numLineas))
                    {
                        row = new Dictionary<string, object>();
                        row.Add("U_SCL_CodeSN", cardCode);
                        rows.Add(row);
                    }
                    JObject jsonObj = new JObject();
                    jsonObj.Add("JournalEntryLines", JArray.FromObject(rows));
                    var json = JsonConvert.SerializeObject(jsonObj);
                    ModifyJournalEntry(json, TransId, sessionID);
                    oRctAsientos.MoveNext();
                    barraProgreso.Value += 1;
                }
                barraProgreso.Stop();
                SBO_Application.StatusBar.SetText("Fueron modificados los asientos ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                GC.Collect();
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("" + ex);
                escribirLog("AsignarTerceroAsientos:" + ex.Message);
            }
        }


        public void ActualizarTipoContAsientos()
        {
            SAPbobsCOM.Recordset oRctCont;
            SAPbouiCOM.ProgressBar barraProgreso;
            try
            {
                oRctCont = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = Properties.Resources.TipoContabilizacion;
                oRctCont.DoQuery(query);
                barraProgreso = SBO_Application.StatusBar.CreateProgressBar("Barra de progreso", oRctCont.RecordCount, false);
                //oConnection.ConCompany(oCompany, SBO_Application);
                string sessionID = oConnection.ConexionServiceLayer();
                while (!oRctCont.EoF)
                {
                    string contDoc = oRctCont.Fields.Item("ContDocumento").Value.ToString();
                    int transId = Convert.ToInt32(oRctCont.Fields.Item("TransId").Value.ToString());
                    var data = new Dictionary<string, object>
                    {
                        {"U_SCL_Contabilizacion", contDoc }
                    };
                    var body = JsonConvert.SerializeObject(data);
                    RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
                    RestRequest request = new RestRequest("JournalEntries(" + transId + ")", Method.PATCH);
                    //Console.WriteLine(cliente.BaseUrl + "" + request.Resource);
                    //Console.WriteLine(body);
                    request.RequestFormat = DataFormat.Json;
                    request.AddCookie("B1SESSION", sessionID);
                    request.AddParameter("application/json", body, ParameterType.RequestBody);
                    RestResponse response = (RestResponse)cliente.Execute(request);
                    //Console.WriteLine(response.Content);
                    oRctCont.MoveNext();
                    barraProgreso.Value += 1;
                    if (response.StatusCode.Equals(HttpStatusCode.BadRequest))
                    {
                        var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                        var jvalue = (JValue)jobject["error"]["message"]["value"];
                        SBO_Application.SetStatusBarMessage(jvalue.Value + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    }
                    else if (response.StatusCode.Equals("301"))
                    {
                        oConnection.ConCompany(oCompany, SBO_Application);
                        sessionID = oConnection.ConexionServiceLayer();
                    }
                }
                barraProgreso.Stop();
                SBO_Application.StatusBar.SetText("Asientos de Notas Credito actualizados", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                GC.Collect();
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox("" + ex);//
                escribirLog("ActualizarTipoContAsientos: " +ex.Message);
            }
        }

        public void AnularNCVentas()
        {
            SAPbobsCOM.Recordset oRecordset;
            SAPbobsCOM.Recordset oRdstLineas;
            SAPbobsCOM.Recordset oRdstGrupArt;
            SAPbobsCOM.Recordset oRdstImpuesto;
            SAPbobsCOM.Recordset oRdstNotaCre;
            SAPbouiCOM.ProgressBar barraProgreso;
            try
            {
                oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRdstLineas = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRdstGrupArt = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRdstImpuesto = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRdstNotaCre = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                decimal valor = 0;
                CancelacionNC cancelarNC = new CancelacionNC(oCompany, SBO_Application);
                string query = string.Format(Consultas.Default.idAsientoCanNC, "U_SCL_DocNCV", "ORIN", "A");
                //query = Properties.Resources.idAsientoCanNC;
                oRecordset.DoQuery(query);
                barraProgreso = SBO_Application.StatusBar.CreateProgressBar("Barra de progreso", oRecordset.RecordCount, false);
                while (!oRecordset.EoF)
                {
                    ArrayList cuentaArt = new ArrayList();
                    ArrayList nombreCnta = new ArrayList();
                    ArrayList valorLinea = new ArrayList();
                    query = Properties.Resources.LinAsientoCanNC.Replace("%", oRecordset.Fields.Item("TransId").Value.ToString());
                    oRdstLineas.DoQuery(query);
                    while (!oRdstLineas.EoF)
                    {
                        string cntAsiento = oRdstLineas.Fields.Item("Account").Value.ToString();
                        query = Properties.Resources.CntasGrupoArtVen;
                        oRdstGrupArt.DoQuery(query);
                        while (!oRdstGrupArt.EoF)
                        {
                            //  string cntGrupoArt = oRdstGrupArt.Fields.Item("CuentaNC").Value.ToString();
                            if (cntAsiento.Equals(oRdstGrupArt.Fields.Item("CuentaNC").Value.ToString()))
                            {
                                valor = Convert.ToDecimal(oRdstLineas.Fields.Item("Debit").Value.ToString());
                                if (valor > 0)
                                {
                                    cuentaArt.Add("A" + oRdstLineas.Fields.Item("Account").Value.ToString());
                                    nombreCnta.Add(oRdstLineas.Fields.Item("AcctName").Value.ToString());
                                    valorLinea.Add(valor);
                                }
                            }
                            oRdstGrupArt.MoveNext();
                        }
                        query = Properties.Resources.CntasIVAVentas;
                        oRdstImpuesto.DoQuery(query);
                        while (!oRdstImpuesto.EoF)
                        {
                            //string cntaIVA = ;
                            if (cntAsiento.Equals(oRdstImpuesto.Fields.Item("SalesTax").Value.ToString()))
                            {
                                valor = Convert.ToDecimal(oRdstLineas.Fields.Item("Debit").Value.ToString());
                                if (valor > 0)
                                {
                                    cuentaArt.Add(oRdstLineas.Fields.Item("Account").Value.ToString());
                                    nombreCnta.Add(oRdstLineas.Fields.Item("AcctName").Value.ToString());
                                    valorLinea.Add(valor);
                                }
                            }
                            oRdstImpuesto.MoveNext();
                        }
                        oRdstLineas.MoveNext();
                        //Crear asiento con los datos recolectados
                    }
                    if (cuentaArt.Count > 0)
                    {
                        query = string.Format(Consultas.Default.DocEntryNC, "ORIN", oRecordset.Fields.Item("TransId").Value.ToString());
                        oRdstNotaCre.DoQuery(query);
                        cancelarNC.generarAsientoVentas(cuentaArt, nombreCnta, valorLinea, Convert.ToInt32(oRdstNotaCre.Fields.Item("DocEntry").Value), Convert.ToInt32(oRecordset.Fields.Item("TransId").Value), oRdstNotaCre.Fields.Item("CardCode").Value);
                    }
                    else
                    {
                        cancelarNC.agregarReferenciaAsiento(Convert.ToInt32(oRecordset.Fields.Item("TransId").Value), Convert.ToInt32(oRecordset.Fields.Item("TransId").Value), 0);
                    }

                    oRecordset.MoveNext();
                    barraProgreso.Value += 1;
                }
                barraProgreso.Stop();
                SBO_Application.StatusBar.SetText("Asientos de Notas credito en ventas actualizados", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                GC.Collect();
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
                GC.Collect();
            }
        }

        public void AnularNCCompras()
        {
            //xxxxxx
            SAPbobsCOM.Recordset oRecordset;
            SAPbobsCOM.Recordset oRdstLineas;
            SAPbobsCOM.Recordset oRdstGrupArt;
            SAPbobsCOM.Recordset oRdstImpuesto;
            SAPbobsCOM.Recordset oRdstNotaCre;
            SAPbouiCOM.ProgressBar barraProgreso;
            try
            {
                oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRdstLineas = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRdstGrupArt = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRdstImpuesto = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRdstNotaCre = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                decimal valor = 0;
                //string sn  = string.Empty;
                CancelacionNC cancelarNC = new CancelacionNC(oCompany, SBO_Application);
                string query = string.Format(Consultas.Default.idAsientoCanNC, "U_SCL_DocNCC", "ORPC", "A");
                //query = Properties.Resources.idAsientoCanNC;
                oRecordset.DoQuery(query);
                barraProgreso = SBO_Application.StatusBar.CreateProgressBar("Barra de progreso", oRecordset.RecordCount, false);
                while (!oRecordset.EoF)
                {
                    ArrayList cuentaArt = new ArrayList();
                    ArrayList nombreCnta = new ArrayList();
                    ArrayList valorLinea = new ArrayList();
                    query = Properties.Resources.LinAsientoCanNC.Replace("%", oRecordset.Fields.Item("TransId").Value.ToString());
                    oRdstLineas.DoQuery(query);
                    while (!oRdstLineas.EoF)
                    {
                        string cntAsiento = oRdstLineas.Fields.Item("Account").Value.ToString();
                        query = Properties.Resources.CntasIVACompras;
                        oRdstImpuesto.DoQuery(query);
                        while (!oRdstImpuesto.EoF)
                        {
                            //string cntaIVA = ;
                            if (cntAsiento.Equals(oRdstImpuesto.Fields.Item("PurchTax").Value.ToString()))
                            {
                                valor = Convert.ToDecimal(oRdstLineas.Fields.Item("Credit").Value.ToString());
                                if (valor > 0)
                                {
                                    cuentaArt.Add(oRdstLineas.Fields.Item("Account").Value.ToString());
                                    nombreCnta.Add(oRdstLineas.Fields.Item("AcctName").Value.ToString());
                                    valorLinea.Add(valor);
                                }
                            }
                            oRdstImpuesto.MoveNext();
                        }
                        oRdstLineas.MoveNext();
                        //Crear asiento con los datos recolectados
                    }
                    // query = Properties.Resources.DocEntryNC.Replace("%", oRecordset.Fields.Item("TransId").Value.ToString());

                    //cancelarNC.generarAsientoVentas(cuentaArt, nombreCnta, valorLinea, Convert.ToInt32(oRdstNotaCre.Fields.Item("DocEntry").Value), Convert.ToInt32(oRecordset.Fields.Item("TransId").Value));
                    if (cuentaArt.Count > 0)
                    {
                        query = string.Format(Consultas.Default.DocEntryNC, "ORPC", oRecordset.Fields.Item("TransId").Value.ToString());
                        oRdstNotaCre.DoQuery(query);
                        cancelarNC.generarAsientoCompras(cuentaArt, nombreCnta, valorLinea, Convert.ToInt32(oRdstNotaCre.Fields.Item("DocEntry").Value), Convert.ToInt32(oRecordset.Fields.Item("TransId").Value), oRdstNotaCre.Fields.Item("CardCode").Value);
                    }
                    else
                    {
                        cancelarNC.agregarReferenciaAsiento(Convert.ToInt32(oRecordset.Fields.Item("TransId").Value), Convert.ToInt32(oRecordset.Fields.Item("TransId").Value), 0);
                    }

                    oRecordset.MoveNext();
                    barraProgreso.Value += 1;
                }
                barraProgreso.Stop();
                SBO_Application.StatusBar.SetText("Asientos de Notas credito en compras actualizados", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                GC.Collect();
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
                GC.Collect();
            }
        }

        //---------------- Crear Documento de Revalorizacion de inventario ------------
        public void addMaterialRevaluationIVAMV(ArrayList articulos, string cnta, string SN, int numDocEM, string sessionID)//kkkkkkkkkkkkkkkkkkkkkkkkkkkkkk
        {               
            List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
            Dictionary<string, object> row;
            SAPbobsCOM.Recordset oRecordset;
            oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                foreach (IVA_Mayor.ArticuloIVA itm in articulos)
                {
                    row = new Dictionary<string, object>();
                    row.Add("ItemCode", itm.codigo); //Codigo articulo
                    row.Add("Quantity", itm.cantidad);//Costos nuevos
                    row.Add("ActualPrice", itm.costoActual);//Costos actuales
                                                            //row.Add("Price", itm.costoNuevo);//Costos nuevos
                    row.Add("WarehouseCode", itm.almacen); //Almacen
                    row.Add("DebitCredit", itm.costoNuevo); //Debito/Credito
                    row.Add("RevaluationDecrementAccount", cnta); //Disminuir cuenta de mayor
                    row.Add("RevaluationIncrementAccount", cnta); //Aumentar cuenta de mayor
                    rows.Add(row);
                }
                JObject jsonObj = new JObject();
                jsonObj.Add("Comments", "Rev. con base en Entrada de Mercancia n.º " + numDocEM);
                jsonObj.Add("U_SCL_CodSN", SN);
                jsonObj.Add("RevalType", "M");
                jsonObj.Add("DataSource", "I");
                jsonObj.Add("MaterialRevaluationLines", JArray.FromObject(rows));
                var body = JsonConvert.SerializeObject(jsonObj);
                RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
                RestRequest request = new RestRequest("MaterialRevaluation", Method.POST);
                request.RequestFormat = DataFormat.Json;
                request.AddCookie("B1SESSION", sessionID);
                request.AddParameter("application/json", body, ParameterType.RequestBody);
                RestResponse response = (RestResponse)cliente.Execute(request);
                var res = response.Content;
                //Console.WriteLine(response.StatusDescription);
                if (response.StatusCode.Equals(HttpStatusCode.Created))
                {
                    dynamic dynJson = JsonConvert.DeserializeObject(res);
                    string TransId = "";
                    string DocEntry = "";
                    string DocNumRI = "";
                    foreach (var item in dynJson)
                    {
                        if (item.Name == "TransNum") TransId = item.Value;
                        if (item.Name == "DocEntry") DocEntry = item.Value;
                        if (item.Name == "DocNum") DocNumRI = item.Value;
                    }

                    string queryLineas = Properties.Resources.NumLineasAsiento;
                    oRecordset.DoQuery(queryLineas.Replace("%", TransId));
                    int numLineas = Convert.ToInt32(oRecordset.Fields.Item("NUMFILAS").Value);
                    rows = new List<Dictionary<string, object>>();
                    foreach (var index in Enumerable.Range(1, numLineas))
                    {
                        row = new Dictionary<string, object>();
                        row.Add("U_SCL_CodeSN", SN);
                        row.Add("U_SCL_TipoMM", 5);
                        rows.Add(row);
                    }
                    jsonObj = new JObject();
                    jsonObj.Add("JournalEntryLines", JArray.FromObject(rows));
                    jsonObj.Add("Memo", "Rev. con base en Entrada de Mercancia n.º " + numDocEM);
                    jsonObj.Add("TransactionCode", "IMVG");
                    jsonObj.Add("Reference", DocNumRI);
                    var json = JsonConvert.SerializeObject(jsonObj);
                    ModifyJournalEntry(json, TransId, sessionID);
                    SBO_Application.OpenForm((SAPbouiCOM.BoFormObjectEnum)162, "", DocEntry.ToString());


                }
                else
                {
                    var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                    var jvalue = (JValue)jobject["error"]["message"]["value"];
                    SBO_Application.SetStatusBarMessage(jvalue.Value + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    escribirLog("MaterialRevaluation: " + jvalue.Value + oCompany.GetLastErrorDescription());
                }

            }
            catch(Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
                GC.Collect();
            }
        }
        //-------------------------------------------------------------------------------------

        void CreateJournalEntryIVAMV(string cnta1, double valor, string cnta2, string SN, int numDoc, string sessionID, int contador)
        {
            List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
            Dictionary<string, object> row;
            //int contador = 1;
            foreach (var index in Enumerable.Range(1, 2))
            {
                if (contador % 2 != 0)
                {
                    row = new Dictionary<string, object>();
                    row.Add("AccountCode", cnta1);
                    row.Add("Credit", valor);
                    row.Add("Debit", 0);
                    row.Add("U_SCL_CodeSN", SN);
                    row.Add("U_SCL_TipoMM", 5);
                    rows.Add(row);
                }else
                {
                    row = new Dictionary<string, object>();
                    row.Add("AccountCode", cnta2);
                    row.Add("Credit", 0);
                    row.Add("Debit", valor);
                    row.Add("U_SCL_CodeSN", SN);
                    row.Add("U_SCL_TipoMM", 5);
                    rows.Add(row);
                }
                contador++;
            }
            JObject jsonObj = new JObject();
            jsonObj.Add("JournalEntryLines", JArray.FromObject(rows));
            jsonObj.Add("Memo", "Con base en Entrada de Mercancia n.º " + numDoc);
            jsonObj.Add("TransactionCode", "IMVG");
            jsonObj.Add("Reference", numDoc);
            var json = JsonConvert.SerializeObject(jsonObj);
            AddJournalEntry(json, sessionID);
        }
        //--------------------------------- CREAR / MODIFICAR ASIENTO ----------------------------------------------------

        void ModifyJournalEntry(string json, string transId, string sessionID)
        {

           try
            {
                RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
                RestRequest request = new RestRequest("JournalEntries(" + transId + ")", Method.PATCH);
                request.AddCookie("B1SESSION", sessionID);
                request.AddParameter("application/json", json, ParameterType.RequestBody);
                RestResponse response = (RestResponse)cliente.Execute(request);
                if (response.StatusCode.Equals(HttpStatusCode.BadRequest))
                {
                    var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                    var jvalue = (JValue)jobject["error"]["message"]["value"];
                    SBO_Application.SetStatusBarMessage("ModifyJournalEntry: " + jvalue.Value + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    escribirLog("ModifyJournalEntry: " + jvalue.Value + oCompany.GetLastErrorDescription());
                }
                else if (response.StatusCode.Equals("301"))
                {
                    oConnection.ConCompany(oCompany, SBO_Application);
                    sessionID = oConnection.ConexionServiceLayer();
                    ModifyJournalEntry(json, transId, sessionID);
                }                
            }
            catch (Exception ex)
            {
                escribirLog("ModifyJournalEntry:" + ex.Message);
            }
        }

        void AddJournalEntry(string json, string sessionID)
        {
            try
            {
                RestClient cliente = new RestClient(DatosGlobServiceLayer.url);
                RestRequest request = new RestRequest("JournalEntries", Method.POST);
                request.AddCookie("B1SESSION", sessionID);
                request.AddParameter("application/json", json, ParameterType.RequestBody);
                RestResponse response = (RestResponse)cliente.Execute(request);
                var res = response.Content;
                if (response.StatusCode.Equals(HttpStatusCode.Created))
                {
                    dynamic dynJson = JsonConvert.DeserializeObject(res);
                    int TransId = 0;
                    foreach (var item in dynJson)
                    {
                        if (item.Name == "JdtNum") TransId = item.Value;
                    }
                    SBO_Application.OpenForm((SAPbouiCOM.BoFormObjectEnum)30, "", TransId.ToString());
                }
                else if (response.StatusCode.Equals(HttpStatusCode.BadRequest))
                {
                    var jobject = (JObject)JsonConvert.DeserializeObject(response.Content);
                    var jvalue = (JValue)jobject["error"]["message"]["value"];
                    SBO_Application.SetStatusBarMessage("AddJournalEntry: " + jvalue.Value + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    escribirLog("AddJournalEntry: " + jvalue.Value + oCompany.GetLastErrorDescription());
                }
                else if (response.StatusCode.Equals("301"))
                {
                    oConnection.ConCompany(oCompany, SBO_Application);
                    sessionID = oConnection.ConexionServiceLayer();
                    AddJournalEntry(json, sessionID);
                }
            }
            catch(Exception ex)
            {
                escribirLog("AddJournalEntry:" + ex.Message);
            }
        }
        //---------------------------------------------------------------------------------------
        //void SBO_Application_LayoutKeyEvent(ref SAPbouiCOM.LayoutKeyInfo eventInfo, out bool BubbleEvent)
        //{
        //    BubbleEvent = true;

        //    if (eventInfo.ReportTemplate == "A001" && eventInfo.ReportCode == "A001001")
        //    {
        //        eventInfo.LayoutKey = ""; //Set the key of the layout 
        //    }
        //}

    }
}