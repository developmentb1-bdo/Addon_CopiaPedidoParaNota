using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace Addon_CopiaPedidoParaNota
{
    class Menu
    {
        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = Application.SBO_Application.Menus.Item("43520"); // moudles'

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "Addon_CopiaPedidoParaNota";
            oCreationPackage.String = "Faturamento em Lote";
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = -1;

            oMenus = oMenuItem.SubMenus;

            try
            {
                //  If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception e)
            {

            }

            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("Addon_CopiaPedidoParaNota");
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "Addon_CopiaPedidoParaNota.Form1";
                oCreationPackage.String = "Criar Nota Fiscal em Lote";
                oMenus.AddEx(oCreationPackage);

                //oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                //oCreationPackage.UniqueID = "Addon_CopiaPedidoParaNota.FormPicking";
                //oCreationPackage.String = "Autorizar e Efetuar Picking";
                //oMenus.AddEx(oCreationPackage);
            }
            catch (Exception er)
            { //  Menu already exists
                //Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "Addon_CopiaPedidoParaNota.Form1")
                {
                    Form1 activeForm = new Form1();
                    activeForm.Show();
                }

                //if (pVal.BeforeAction && pVal.MenuUID == "Addon_CopiaPedidoParaNota.FormPicking")
                //{
                //    FormPicking activeForm = new FormPicking();
                //    activeForm.Show();
                //}
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

    }
}
