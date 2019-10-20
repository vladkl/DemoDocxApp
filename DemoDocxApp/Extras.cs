/// <summary>
/// Useful new features in Cradle 7.0
/// </summary>
using Cradle.Lists;
using Cradle.HelpSystem;
using Cradle.Data;
using Cradle.Tools;
using Cradle.Server;
using System.Windows.Forms;
using Cradle.Definitions;
using System.Collections.Generic;
using Cradle.ProjectSchema;
using System;

namespace ConsoleApp
{
    static class Example_Extras
    {
        public static Connection conn;

        /// <summary>
        /// This example shows how you can reuse the standard Cradle login dialog
        /// </summary>
        public static void LoginDialog(string bl)
        {
            try
            {
                Globals.Load_CradleAPI();
                conn = new Connection(Globals.CRADLE_CDS_HOST);
                conn.Username = "MANAGER";
                conn.Password = "MANAGER";
                conn.IsReadOnly = false;
                System.Drawing.Icon iconForFile = System.Drawing.SystemIcons.Application;
                conn.CreateLoginDialog("Проверка валидности", iconForFile, Startup, Shutdown, CRADLE_HELP.CRADLE_INDEX, null);
                conn.ShowLoginDialog();
                conn.Logout();
                conn.Login("TBL1", "MANAGER","MANAGER", false);

                Project proj = conn.ActiveProject;
                conn.IsReadOnly = false;
                proj.SetBaselineMode(CAPI_BASELINE_MODE.SPECIFIED, bl);

            }
            catch (CAPIException ex)
            {
                MessageBox.Show(ex.Message, "Cradle API Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                if (ex.Message == "Already connected.") { Shutdown(); }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        /// <summary>
        /// This example shows how you can reuse the standard Cradle definition chooser dialog.
        /// </summary>
        /// <remarks></remarks>
        public static void DefinitionChooser(string bl, List<Item> item_list, string ElemType)
        {
            Project proj = conn.ActiveProject;
            LDAPInformation ldap = null;
            Definition defn = new Definition(CAPI_DEFN_TYPE.QUERY);
            Query query = null;

            //List results = new List(CAPI_LIST.IHIST_ATTRS);
            List results = new List(CAPI_LIST.ITEMS);
            //List<ItemHistoryAttribute> hlist  = new  List<ItemHistoryAttribute>();

            try
            {
                // Connect to the CDS
                //Globals.Load_CradleAPI();
                //proj = new Project();
                //ldap = new LDAPInformation();
                //if (!proj.Connect(Globals.CRADLE_CDS_HOST, Globals.CRADLE_PROJECT_CODE, Globals.CRADLE_USERNAME, Globals.CRADLE_PASSWORD, false, Connection.API_LICENCE, ldap, false))
                //    return;

                // Show the definition chooser dialog
                //if (defn.Choose(null, Definition.Operation.Create))

                query = new Query();
                if (!query.Load(ElemType + " - All", CAPI_DEFN_LOC.AUTO))
                    return;
                proj.SetBaselineMode(CAPI_BASELINE_MODE.SPECIFIED, bl);
                if (query.Run(1000, out results))
                {
                    results.ToDotNetList(ref item_list);
                    results.Dispose();
                }
                query.Dispose();

            }
            catch (CAPIException ex)
            {
                MessageBox.Show(ex.Message, "Cradle API Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                ////  Disconnect from CDS
                // if (proj != null && proj.IsConnected)
                //    proj.Disconnect();
            }
        }

        private static void Startup()
        {
            MessageBox.Show("Program starting...");
        }

        private static void Shutdown()
        {
            if (conn != null && conn.IsActive)
                conn.Dispose();
            MessageBox.Show("Program exiting...");
        }
    }
}
