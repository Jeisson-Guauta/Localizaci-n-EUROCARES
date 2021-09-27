using System;
using System.Xml;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;
using System.Globalization;

namespace LocalizacionColombia
{
    public class FuncionesGenerales
    {
        public static SAPbouiCOM.Form CargarFormularioXML(SAPbouiCOM.Application oApp, string strRuta, string strPrefijo)
        {
            #region Variables Locales
            Form oFrm = null;
            FormCreationParams oFormCreationParams = null;
            XmlDocument oXmlDoc = null;
            string strXML = string.Empty, strError = string.Empty;
            #endregion

            try
            {
                oFormCreationParams = (FormCreationParams)oApp.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
                oXmlDoc = new XmlDocument();
                oXmlDoc.Load(strRuta);
                strXML = oXmlDoc.InnerXml.ToString();

                oFormCreationParams.XmlData = strXML;
                oFormCreationParams.UniqueID = generarIDForm(oApp, strPrefijo);

                oFrm = oApp.Forms.AddEx(oFormCreationParams);
                oFrm.Visible = true; 
                strError = oApp.GetLastBatchResults();

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return oFrm;
        }

        public static string generarIDForm(SAPbouiCOM.Application oApp, string strPrefijo)
        {
            #region Variables Locales
            string strIdForm = string.Empty;
            SAPbouiCOM.Form oFrm = null;
            #endregion

            for (int i = 1; i < 10; i++)
            {
                try
                {
                    oFrm = oApp.Forms.Item(strPrefijo + "_0" + i.ToString());
                }
                catch
                {
                    strIdForm = strPrefijo + "_0" + i.ToString();
                    break;
                }
                finally
                {
                    if (oFrm != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oFrm);
                    }
                }
            }

            return strIdForm;
        }

        public void Set_Conditions(SAPbouiCOM.Application oApplication, string Alias, string Alias2, string Value, string Value2, SAPbouiCOM.DBDataSource oDBDataSource)
        {
            try
            {
                SAPbouiCOM.Conditions oCons;
                SAPbouiCOM.Condition oCon;
                oCons = (SAPbouiCOM.Conditions)oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                oCon = oCons.Add();
                oCon.BracketOpenNum = 2;
                oCon.Alias = Alias;
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = Value;
                oCon.BracketCloseNum = 1;
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCon = oCons.Add();
                oCon.BracketOpenNum = 1;
                oCon.Alias = Alias2;
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = Value2;
                oCon.BracketCloseNum = 2;
                oCon.BracketCloseNum = 2;
                // Querying the DB Data source
                oDBDataSource.Query(oCons);
            }
            catch { }
        }

        public static void CargarSeriesUDO(SAPbobsCOM.Company oCmp, SAPbouiCOM.Form oFrm, string strUDO, string strCombo, string strDocNum)
        {
            #region Variables Locales
            SAPbobsCOM.Recordset oRS = null;
            SAPbouiCOM.ComboBox oCmb = null;
            string strQry = string.Empty;
            #endregion

            try
            {
                oCmb = (SAPbouiCOM.ComboBox)oFrm.Items.Item(strCombo).Specific;
                oRS = (SAPbobsCOM.Recordset)oCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                oCmb.ValidValues.LoadSeries(strUDO, BoSeriesMode.sf_Add);

                if (oCmb.ValidValues.Count > 0)
                {
                    strQry = "SELECT \"DfltSeries\" FROM ONNM WHERE \"ObjectCode\"='" + strUDO + "'";
                    oRS.DoQuery(strQry);
                    if (oRS.RecordCount > 0)
                    {
                        oCmb.Select(oRS.Fields.Item(0).Value.ToString(), BoSearchKey.psk_ByValue);
                    }
                    else
                    {
                        oCmb.Select(0, BoSearchKey.psk_Index);
                    }
                    ((SAPbouiCOM.EditText)oFrm.Items.Item(strDocNum).Specific).Value = oFrm.BusinessObject.GetNextSerialNumber(oCmb.Selected.Value, strUDO).ToString();
                }
            }
            catch (Exception ex) { }
        }

        public static double TRM(SAPbobsCOM.Company oCmp)
        {
            #region Variables Locales
            SAPbobsCOM.Recordset oRS = null;
            double dbTRM = 0;
            string strQry = string.Empty;
            #endregion

            try
            {
                oRS = (SAPbobsCOM.Recordset)oCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                strQry = "SELECT CAST(\"Rate\" AS decimal(6,2)) AS \"Rate\" FROM ORTT WHERE \"RateDate\" = CURRENT_DATE AND \"Currency\" = 'USD'";
                oRS.DoQuery(strQry);
                if (oRS.RecordCount > 0)
                {
                    dbTRM = Convert.ToDouble(oRS.Fields.Item(0).Value.ToString());
                }
            }
            catch { }

            return dbTRM;
        }

        public static double PorcentajeIVA(SAPbobsCOM.Company oCmp, string strCodIva)
        {
            #region Variables Locales
            SAPbobsCOM.Recordset oRS = null;
            double dbPorcIva = 0;
            string strQry = string.Empty;
            #endregion

            try
            {
                if (!strCodIva.Equals(""))
                {
                    oRS = (SAPbobsCOM.Recordset)oCmp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    strQry = "SELECT \"Code\", \"Name\", \"Rate\" FROM OSTA WHERE IFNULL(\"PurchTax\", '') != '' AND \"Code\" = '" + strCodIva + "'";
                    oRS.DoQuery(strQry);
                    if (oRS.RecordCount > 0)
                    {
                        dbPorcIva = Convert.ToDouble(oRS.Fields.Item(2).Value.ToString());
                    }
                }
            }
            catch { }

            return dbPorcIva;
        }

        public static double StringADoubleBD(SAPbobsCOM.Company oCmp, string strValor)
        {
            #region Variables Locales
            double dValor = 0;
            CultureInfo confSBOI = new CultureInfo("es-CO", false);
            CultureInfo confDB = new CultureInfo("es-CO", false);
            //string strValor;
            #endregion
            try
            {
                #region CultureInfo


                SAPbobsCOM.CompanyService oCompanyService;
                SAPbobsCOM.AdminInfo oAdminInfo;
                oCompanyService = (SAPbobsCOM.CompanyService)oCmp.GetCompanyService();
                oAdminInfo = oCompanyService.GetAdminInfo();

                confSBOI.NumberFormat.NumberDecimalSeparator = oAdminInfo.DecimalSeparator;
                confSBOI.NumberFormat.NumberGroupSeparator = oAdminInfo.ThousandsSeparator;

                confDB.NumberFormat.NumberDecimalSeparator = ".";
                confDB.NumberFormat.NumberGroupSeparator = ",";
                #endregion
                //Lo traigo con la configuracion de SBO
                //strValor = Convert.ToString(objValor, confSBOI);
                strValor = strValor.Replace(confDB.NumberFormat.NumberGroupSeparator, "");
                //Lo entrego con la configuracion de la Base de Datos .
                if (strValor.Contains("."))
                {
                    dValor = double.Parse(strValor, confDB);
                }
                else
                {
                    dValor = double.Parse(strValor, confSBOI);
                }
            }
            catch (Exception e)
            {
                try
                {
                    //strValor = Convert.ToString(objValor, confDB);
                    strValor = strValor.Replace(confDB.NumberFormat.NumberGroupSeparator, "");

                    //Lo entrego con la configuracion de la Base de Datos .
                    dValor = double.Parse(strValor, confDB);
                }
                catch { }
                //throw e; 

            }
            return dValor;
        }

        public static double StringADoubleForm(SAPbobsCOM.Company oCmp, string strValor)
        {
            #region Variables Locales
            double dValor = 0;
            CultureInfo confSBOI = new CultureInfo("es-CO", false);
            CultureInfo confDB = new CultureInfo("es-CO", false);
            //string strValor;
            #endregion
            try
            {
                #region CultureInfo


                SAPbobsCOM.CompanyService oCompanyService;
                SAPbobsCOM.AdminInfo oAdminInfo;
                oCompanyService = (SAPbobsCOM.CompanyService)oCmp.GetCompanyService();
                oAdminInfo = oCompanyService.GetAdminInfo();

                confSBOI.NumberFormat.NumberDecimalSeparator = oAdminInfo.DecimalSeparator;
                confSBOI.NumberFormat.NumberGroupSeparator = oAdminInfo.ThousandsSeparator;

                confDB.NumberFormat.NumberDecimalSeparator = ".";
                confDB.NumberFormat.NumberGroupSeparator = ",";
                #endregion
                //Lo traigo con la configuracion de SBO
                //strValor = Convert.ToString(objValor, confSBOI);
                strValor = strValor.Replace(confSBOI.NumberFormat.NumberGroupSeparator, "");
                //Lo entrego con la configuracion de la Base de Datos .
                if (strValor.Contains("."))
                {
                    dValor = double.Parse(strValor, confDB);
                }
                else
                {
                    dValor = double.Parse(strValor, confSBOI);
                }
            }
            catch (Exception e)
            {
                try
                {
                    //strValor = Convert.ToString(objValor, confDB);
                    strValor = strValor.Replace(confDB.NumberFormat.NumberGroupSeparator, "");

                    //Lo entrego con la configuracion de la Base de Datos .
                    dValor = double.Parse(strValor, confDB);
                }
                catch { }
                //throw e; 

            }
            return dValor;
        }

        public static string DoubleAString(SAPbobsCOM.Company oCmp, object objValor)
        {
            #region Variables locales
            string strValor = string.Empty;
            double dbValor = 0;
            #endregion
            try
            {
                #region CultureInfo
                CultureInfo confSBOI = new CultureInfo("es-CO", false);
                CultureInfo confDB = new CultureInfo("es-CO", false);

                SAPbobsCOM.CompanyService oCompanyService;
                SAPbobsCOM.AdminInfo oAdminInfo;
                int DecimalesVis = 0;
                try
                {
                    oCompanyService = (SAPbobsCOM.CompanyService)oCmp.GetCompanyService();
                    oAdminInfo = oCompanyService.GetAdminInfo();
                    DecimalesVis = oAdminInfo.PercentageAccuracy;

                    confSBOI.NumberFormat.NumberDecimalSeparator = oAdminInfo.DecimalSeparator;
                    confSBOI.NumberFormat.NumberGroupSeparator = oAdminInfo.ThousandsSeparator;

                    confDB.NumberFormat.NumberDecimalSeparator = ".";
                    confDB.NumberFormat.NumberGroupSeparator = ",";
                }
                catch (Exception e)
                { throw e; }
                #endregion

                try
                {
                    //Traigo el valor con referencia a .
                    dbValor = Convert.ToDouble(objValor, confDB);
                    //lo redondeo al numero de decimales
                    dbValor = Math.Round(dbValor, DecimalesVis);
                    //lo entrego con base en el .
                    strValor = Convert.ToString(dbValor, confSBOI);
                }
                catch (Exception e)
                { throw e; }
            }
            catch (Exception e)
            { throw e; }

            return strValor;
        }
    }
}
