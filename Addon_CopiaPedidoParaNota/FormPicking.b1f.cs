using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Addon_CopiaPedidoParaNota
{
    [FormAttribute("Addon_CopiaPedidoParaNota.FormPicking", "FormPicking.b1f")]
    internal class FormPicking : UserFormBase
    {
        public FormPicking()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_2").Specific));
            this.Button1.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button1_ClickAfter);
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_3").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_4").Specific));
            this.Button2.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button2_ClickAfter);
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("Item_5").Specific));
            this.Button3.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button3_ClickAfter);
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("Item_6").Specific));
            this.Button4.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button4_ClickAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private void OnCustomInitialize()
        {
            SAPbouiCOM.Form oForm = (SAPbouiCOM.Form)this.UIAPIRawForm;
            Conexao.CentralizarForm(oForm);
        }

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button Button3;
        private SAPbouiCOM.Button Button4;

        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form form = Conexao.sbo_application.Forms.Item((object)pVal.FormUID);

            try
            {
                form.Freeze(true);

                Conexao.sbo_application.StatusBar.SetText($@"Consultando os Pedidos de Venda de RS."
                               , BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                string query = $@"SELECT 
                                    'N' ""Sel."", 
                                    T0.""DocEntry"" ""Nº Primário"", 
                                    T0.""DocNum"" ""Nº do Doc"", 
                                    T0.""DocDate"" ""Data de Lançamento"", 
                                    T0.""CardCode"" ""Cód. PN"", 
                                    T0.""CardName"" ""Nome PN"", 
                                    T2.""State"" ""UF do Cliente"", 
                                    T0.""DocTotal"" ""Valor Total""
                                FROM ORDR T0 
                                    INNER JOIN OCRD T1 ON T0.""CardCode"" = T1.""CardCode"" 
                                    INNER JOIN CRD1 T2 ON T1.""CardCode"" = T2.""CardCode"" 
                                WHERE 
                                    T0.""Pick"" = 'N' 
                                    AND T0.""Confirmed"" = 'N' 
                                    AND T2.""State"" = 'RS'
                                    AND T0.""CANCELED"" = 'N'
                                    AND T0.""DocStatus"" = 'O'
                                    AND T2.""AdresType"" = 'S'
                                ORDER BY T0.""DocEntry""";

                Grid0.DataTable.ExecuteQuery(query);

                GridColumn oCol = Grid0.Columns.Item(0);
                oCol.Type = BoGridColumnType.gct_CheckBox;

                EditTextColumn oColDocEntry = (EditTextColumn)Grid0.Columns.Item(1);
                oColDocEntry.LinkedObjectType = "17";

                Grid0.Columns.Item(2).Editable = false;
                Grid0.Columns.Item(3).Editable = false;
                Grid0.Columns.Item(4).Editable = false;
                Grid0.Columns.Item(5).Editable = false;
                Grid0.Columns.Item(6).Editable = false;
                Grid0.Columns.Item(7).Editable = false;

                Button1.Item.Enabled = true;
                Button2.Item.Enabled = true;
                Button3.Item.Enabled = true;

                Conexao.sbo_application.StatusBar.SetText($@"Consulta dos Pedidos de RS realizada com sucesso."
                                , BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Conexao.sbo_application.StatusBar.SetText($@"Erro para listar os Pedido de Venda [AEP00001]: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                form.Freeze(false);
            }
        }

        private void Button1_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form form = Conexao.sbo_application.Forms.Item((object)pVal.FormUID);

            try
            {
                form.Freeze(true);

                Conexao.sbo_application.StatusBar.SetText($@"Consultando os Pedidos de Venda de outros estados."
                               , BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                string query = $@"SELECT 
                                    'N' ""Sel."", 
                                    T0.""DocEntry"" ""Nº Primário"", 
                                    T0.""DocNum"" ""Nº do Doc"", 
                                    T0.""DocDate"" ""Data de Lançamento"", 
                                    T0.""CardCode"" ""Cód. PN"", 
                                    T0.""CardName"" ""Nome PN"", 
                                    T2.""State"" ""UF do Cliente"", 
                                    T0.""DocTotal"" ""Valor Total""
                                FROM ORDR T0 
                                    INNER JOIN OCRD T1 ON T0.""CardCode"" = T1.""CardCode"" 
                                    INNER JOIN CRD1 T2 ON T1.""CardCode"" = T2.""CardCode"" 
                                WHERE 
                                    T0.""Pick"" = 'N' 
                                    AND T0.""Confirmed"" = 'N' 
                                    AND T2.""State"" <> 'RS'
                                    AND T0.""CANCELED"" = 'N'
                                    AND T0.""DocStatus"" = 'O'
                                    AND T2.""AdresType"" = 'S'
                                ORDER BY T0.""DocEntry""";

                Grid0.DataTable.ExecuteQuery(query);

                GridColumn oCol = Grid0.Columns.Item(0);
                oCol.Type = BoGridColumnType.gct_CheckBox;

                EditTextColumn oColDocEntry = (EditTextColumn)Grid0.Columns.Item(1);
                oColDocEntry.LinkedObjectType = "17";

                Grid0.Columns.Item(2).Editable = false;
                Grid0.Columns.Item(3).Editable = false;
                Grid0.Columns.Item(4).Editable = false;
                Grid0.Columns.Item(5).Editable = false;
                Grid0.Columns.Item(6).Editable = false;
                Grid0.Columns.Item(7).Editable = false;

                Button1.Item.Enabled = true;
                Button2.Item.Enabled = true;
                Button3.Item.Enabled = true;

                Conexao.sbo_application.StatusBar.SetText($@"Consulta dos Pedidos outros estados realizada com sucesso."
                                , BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Conexao.sbo_application.StatusBar.SetText($@"Erro para listar os Pedido de Venda [AEP00002]: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                form.Freeze(false);
            }
        }

        private void Button3_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form form = Conexao.sbo_application.Forms.Item((object)pVal.FormUID);

            try
            {
                form.Freeze(true);

                Conexao.sbo_application.StatusBar.SetText($@"Selecionando todas as linhas."
                               , BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                for (int k = 0; k < Grid0.Rows.Count; k++)
                {
                    Grid0.DataTable.SetValue("Sel.", k, "Y");
                }

                Conexao.sbo_application.StatusBar.SetText($@"Linhas Selecionadas com sucesso."
                               , BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Conexao.sbo_application.StatusBar.SetText($@"Erro para Gerar as Notas [AEP00003]: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                form.Freeze(false);
            }
        }

        private void Button4_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form form = Conexao.sbo_application.Forms.Item((object)pVal.FormUID);

            try
            {
                form.Freeze(true);

                Conexao.sbo_application.StatusBar.SetText($@"Desmarcando todas as linhas."
                               , BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                for (int k = 0; k < Grid0.Rows.Count; k++)
                {
                    Grid0.DataTable.SetValue("Sel.", k, "N");
                }

                Conexao.sbo_application.StatusBar.SetText($@"Linhas desmarcadas com sucesso."
                               , BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Conexao.sbo_application.StatusBar.SetText($@"Erro ao desmarcar linhas [AEP00004]: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                form.Freeze(false);
            }
        }

        private void Button2_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form form = Conexao.sbo_application.Forms.Item((object)pVal.FormUID);

            try
            {
                form.Freeze(true);
                int a = 0;
                for (int k = 0; k < Grid0.Rows.Count; k++)
                {
                    string isSelected = Grid0.DataTable.GetValue("Sel.", k).ToString();

                    if (isSelected == "Y")
                    {
                        int DocEntry = Convert.ToInt32(Grid0.DataTable.GetValue("Nº Primário", k));

                        Conexao.sbo_application.StatusBar.SetText($@"Autorizando e Efetuando Picking do Pedido {Grid0.DataTable.GetValue("Nº do Doc", k)}"
                                , BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                        SAPbobsCOM.Documents ORDR = (SAPbobsCOM.Documents)Conexao.diCompany.GetBusinessObject(BoObjectTypes.oOrders);

                        ORDR.GetByKey(DocEntry);

                        ORDR.Confirmed = SAPbobsCOM.BoYesNoEnum.tYES;
                        ORDR.Pick = SAPbobsCOM.BoYesNoEnum.tYES;

                        int ret1 = ORDR.Update();
                        string sErro1 = string.Empty;

                        if (ret1 != 0)
                        {
                            Conexao.diCompany.GetLastError(out ret1, out sErro1);
                            Conexao.sbo_application.StatusBar.SetText($@"Erro para Autorizar e Efetuar Picking do Pedido [AEF00005]: {sErro1}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        }
                        else
                        {
                            Conexao.sbo_application.StatusBar.SetText($@"Autorizar e Efetuar Picking do Pedido {Grid0.DataTable.GetValue("Nº do Doc", k)} com sucesso."
                                , BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                        }
                        a++;
                    }
                }

                if (a == 0)
                {
                    MessageBox.Show("Nenhum Pedido de Venda selecionado, favor selecionar pelo menos um para gerar a Nota.", "Aviso Autorizar e Efetuar Picking");
                }
            }
            catch (Exception ex)
            {
                Conexao.sbo_application.StatusBar.SetText($@"Erro para Gerar as Notas [AEF00006]: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                form.Freeze(false);
            }
        }
    }
}
