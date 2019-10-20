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

namespace CostService
{
    class Login
    {
        public static Connection conn;
        public static void LoginDialog()
        {
            try
            {
                Globals.Load_CradleAPI();
                conn = new Connection(Globals.CRADLE_CDS_HOST);
                conn.Username = "MANAGER";
                conn.Password = "MANAGER";
               // conn.IsReadOnly = false;
                System.Drawing.Icon iconForFile = System.Drawing.SystemIcons.WinLogo;
                conn.CreateLoginDialog("Проверка и вычисление стоимости", iconForFile, Startup, Shutdown, CRADLE_HELP.CRADLE_INDEX, null);
                conn.ShowLoginDialog();
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
        public static void DefinitionChooser(
            string bl1,
            //string bl2,
            List<Item> item_list1
            //, List<Item> item_list2
            )
        {
            Project proj = null;
            LDAPInformation ldap = null;
            Definition defn = new Definition(CAPI_DEFN_TYPE.QUERY);
            Query query = null;
            string qname;

            //List results = new List(CAPI_LIST.IHIST_ATTRS);
            List results = new List(CAPI_LIST.ITEMS);
            //List<ItemHistoryAttribute> hlist  = new  List<ItemHistoryAttribute>();

            try
            {
                //Connect to the CDS
                Globals.Load_CradleAPI();
                proj = new Project();
                ldap = new LDAPInformation();
                if (!proj.Connect(Globals.CRADLE_CDS_HOST, Globals.CRADLE_PROJECT_CODE, Globals.CRADLE_USERNAME, Globals.CRADLE_PASSWORD, true, Connection.API_LICENCE, ldap, false))
                    return;
                //proj = conn.ActiveProject;
                // Show the definition chooser dialog
                //if (defn.Choose(null, Definition.Operation.Create))
                qname = Globals.GetQuerys()[0];
                query = new Query();
                if (!query.Load(qname, CAPI_DEFN_LOC.AUTO))
                    { return; }
                    else
                    {
                        proj.SetBaselineMode(CAPI_BASELINE_MODE.SPECIFIED, bl1);
                        if (query.Run(5000, out results))
                        {
                            results.ToDotNetList(ref item_list1);
                            results.Dispose();
                        }
                        //proj.SetBaselineMode(CAPI_BASELINE_MODE.UNSET, bl2);
                        //if (query.Run(1000, out results))
                        //{
                        //    results.ToDotNetList(ref item_list2);
                        //    results.Dispose();
                        //}

                        query.Dispose();
                    }
              
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
                // Disconnect from CDS
                if (proj != null && proj.IsConnected)
                    proj.Disconnect();
            }
        }
    }
}
