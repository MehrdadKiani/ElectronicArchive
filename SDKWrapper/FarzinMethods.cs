using ICAN.FarzinSDK.WebServices.Proxy;
using System;
using System.Collections.Generic;
using System.Data;
using System.Net;
using System.Text;
using System.Web.Services;

namespace SDKWrapper
{
    #region [Description]
    /// <summary>
    /// Summary description for FarzinServices
    /// </summary>
    [WebService(Namespace = "http://www.ICAN.ir/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    #endregion

    public class FarzinMethods : WebService
    {
        string errorMsg = null;
        public List<string> errorList = new List<string> { };



        [WebMethod]
        public void Insert(List<string> DocsStructure, string UserName, string UserHashPassword)
        {
            try
            {
                Authentication AuthenticationClass = new Authentication();
                eFormManagment eFormManagmentClass = new eFormManagment();
                CookieContainer cookiecontainer = new CookieContainer();
                AuthenticationClass.CookieContainer = cookiecontainer;
                if (AuthenticationClass.Login(UserName, UserHashPassword, out errorMsg))
                {
                    eFormManagmentClass.CookieContainer = cookiecontainer;
                    foreach (var XML in DocsStructure)
                    {
                        errorList.Add(2 + DocsStructure.IndexOf(XML) + " " + eFormManagmentClass.InsertDocument(XML));
                    }
                }
                errorList.Add(errorMsg);
            }
            catch (Exception ex)
            {
                errorList.Add(ex.Message + ex.StackTrace);
            }
        }

        public string BuildDocStructure(string FormEnglishName, string FormTableName, string CreatorUserID, string CreatorRoleID, DataTable FieldsTable)
        {
            StringBuilder xmlStructure = new StringBuilder();
            xmlStructure.Insert(0, "<Document><SourceSoftware name='" + FormEnglishName + "' version='1.0' repository='" + FormTableName + "' /><Structure><BaseFields><Field name='CreatorID' type='int'>" + CreatorUserID + "</Field><Field name='CreatorRoleID' type='int'>" + CreatorRoleID + "</Field></BaseFields>");
            foreach (DataRow row in FieldsTable.Rows)//manipulate fields
            {
                xmlStructure.Append("<Field name='" + row[0] + "' type='" + row[1] + "'>" + row[2] + "</Field>");
            }
            xmlStructure.Append("</Structure></Document>");

            return xmlStructure.ToString();
        }
    }
}
