using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml;

namespace Addon_CopiaPedidoParaNota
{
    [FormAttribute("Addon_CopiaPedidoParaNota.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {
        }

        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_1").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_2").Specific));
            this.Button1.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button1_ClickAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_3").Specific));
            this.Button2.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button2_ClickAfter);
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("Item_4").Specific));
            this.Button3.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button3_ClickAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_6").Specific));
            this.OnCustomInitialize();
        }

        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button Button3;
        private EditText EditText0;

        private void OnCustomInitialize()
        {
            SAPbouiCOM.Form oForm = (SAPbouiCOM.Form)this.UIAPIRawForm;
            Conexao.CentralizarForm(oForm);
        }

        private void Button0_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form form = Conexao.sbo_application.Forms.Item((object)pVal.FormUID);

            try
            {
                if (string.IsNullOrEmpty(EditText0.Value.ToString()))
                {
                    MessageBox.Show("Campo Ref. Viagem TMS está vazio, favor preencher esse campo.", "Aviso Addon Faturamento em Lote");
                }
                else
                {
                    form.Freeze(true);

                    Conexao.sbo_application.StatusBar.SetText($@"Consultando os Pedidos de Venda."
                                   , BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                    string query = $@"SELECT 'N' ""Sel."", 
                                         ""DocEntry"" ""Nº Primário"", 
                                         ""DocNum"" ""Nº do Doc"", 
                                         ""CardCode"" ""Cod. PN"", 
                                         ""CardName"" ""Nome do PN"" 
                                  FROM ORDR
                                  WHERE ""CANCELED"" = 'N' 
                                  AND ""DocStatus"" = 'O'
                                  AND ""U_ref_viagem_tms"" like '%{EditText0.Value}'";

                    Grid0.DataTable.ExecuteQuery(query);

                    GridColumn oCol = Grid0.Columns.Item(0);
                    oCol.Type = BoGridColumnType.gct_CheckBox;

                    EditTextColumn oColDocEntry = (EditTextColumn)Grid0.Columns.Item(1);
                    oColDocEntry.LinkedObjectType = "17";

                    Grid0.Columns.Item(2).Editable = false;
                    Grid0.Columns.Item(3).Editable = false;
                    Grid0.Columns.Item(4).Editable = false;

                    Button1.Item.Enabled = true;
                    Button2.Item.Enabled = true;
                    Button3.Item.Enabled = true;

                    Conexao.sbo_application.StatusBar.SetText($@"Consulta dos Pedidos realizada com sucesso."
                                    , BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }

                
            }
            catch (Exception ex)
            {
                Conexao.sbo_application.StatusBar.SetText($@"Erro para listar os Pedido de Venda [PV00001]: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                int a = 0;
                for (int k = 0; k < Grid0.Rows.Count; k++)
                {
                    string isSelected = Grid0.DataTable.GetValue("Sel.", k).ToString();

                    if (isSelected == "Y")
                    {
                        int DocEntry = Convert.ToInt32(Grid0.DataTable.GetValue("Nº Primário", k));

                        Conexao.sbo_application.StatusBar.SetText($@"Gerar Nota Fiscal do Pedido {Grid0.DataTable.GetValue("Nº do Doc", k)}"
                                , BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                        SAPbobsCOM.Documents ORDR = (SAPbobsCOM.Documents)Conexao.diCompany.GetBusinessObject(BoObjectTypes.oOrders);

                        ORDR.GetByKey(DocEntry);

                        SAPbobsCOM.Documents OINV = (SAPbobsCOM.Documents)Conexao.diCompany.GetBusinessObject(BoObjectTypes.oInvoices);

                        OINV.CardCode = ORDR.CardCode;
                        OINV.CardName = ORDR.CardName;
                        OINV.ContactPersonCode = ORDR.ContactPersonCode;
                        OINV.NumAtCard = ORDR.NumAtCard;
                        OINV.BPL_IDAssignedToInvoice = ORDR.BPL_IDAssignedToInvoice;
                        OINV.DocDate = ORDR.DocDate;
                        OINV.DocDueDate = ORDR.DocDueDate;
                        OINV.TaxDate = ORDR.TaxDate;
                        OINV.SalesPersonCode = ORDR.SalesPersonCode;
                        //OINV.DocumentsOwner = ORDR.DocumentsOwner == 0 ? -1 : ORDR.DocumentsOwner;
                        OINV.DocTotal = ORDR.DocTotal;
                        OINV.Comments = ORDR.Comments;
                        //OINV.Series = ORDR.Series;
                        //OINV.SeriesString = ORDR.SeriesString;
                        //OINV.SequenceCode = ORDR.SequenceCode;
                        //OINV.SequenceModel = ORDR.SequenceModel;
                        //OINV.SequenceSerial = ORDR.SequenceSerial;
                        OINV.JournalMemo = ORDR.JournalMemo;
                        OINV.ClosingRemarks = ORDR.ClosingRemarks;
                        OINV.OpeningRemarks = ORDR.OpeningRemarks;
                        OINV.PaymentGroupCode = ORDR.PaymentGroupCode;
                        OINV.PaymentMethod = ORDR.PaymentMethod;
                        OINV.Address = ORDR.Address;
                        OINV.Address2 = ORDR.Address2;
                        OINV.PayToCode = ORDR.PayToCode;
                        OINV.ShipToCode = ORDR.ShipToCode;
                        OINV.UserFields.Fields.Item("U_WB_RouteNumber").Value = ORDR.UserFields.Fields.Item("U_ref_viagem_tms").Value;

                        for (int i = 0; i < ORDR.Lines.Count; i++)
                        {
                            ORDR.Lines.SetCurrentLine(i);

                            OINV.Lines.ItemCode = ORDR.Lines.ItemCode;
                            OINV.Lines.ItemDescription = ORDR.Lines.ItemDescription;
                            OINV.Lines.Quantity = ORDR.Lines.Quantity;
                            OINV.Lines.UnitPrice = ORDR.Lines.UnitPrice;
                            OINV.Lines.Price = ORDR.Lines.Price;
                            OINV.Lines.Usage = ORDR.Lines.Usage;
                            OINV.Lines.TaxCode = ORDR.Lines.TaxCode;
                            OINV.Lines.BaseEntry = ORDR.DocEntry;
                            OINV.Lines.BaseLine = ORDR.Lines.LineNum;
                            OINV.Lines.BaseType = 17;
                            OINV.Lines.CFOPCode = ORDR.Lines.CFOPCode;
                            OINV.Lines.WarehouseCode = ORDR.Lines.WarehouseCode;
                            OINV.Lines.CostingCode = ORDR.Lines.CostingCode;
                            OINV.Lines.CostingCode2 = ORDR.Lines.CostingCode2;
                            OINV.Lines.CostingCode3 = ORDR.Lines.CostingCode3;
                            OINV.Lines.CostingCode4 = ORDR.Lines.CostingCode4;
                            OINV.Lines.CostingCode5 = ORDR.Lines.CostingCode5;
                            OINV.Lines.ProjectCode = ORDR.Lines.ProjectCode;
                            OINV.Lines.AccountCode = ORDR.Lines.AccountCode;
                            OINV.Lines.UseBaseUnits = ORDR.Lines.UseBaseUnits;
                            OINV.Lines.UnitsOfMeasurment = ORDR.Lines.UnitsOfMeasurment;


                            for (int LineNum3 = 0; LineNum3 < ORDR.Lines.BatchNumbers.Count; ++LineNum3)
                            {
                                BatchNumbers batchNumbers1 = ORDR.Lines.BatchNumbers;
                                
                                batchNumbers1.SetCurrentLine(LineNum3);

                                if (!string.IsNullOrEmpty(batchNumbers1.BatchNumber))
                                {
                                    OINV.Lines.BatchNumbers.BatchNumber = batchNumbers1.BatchNumber;
                                    OINV.Lines.BatchNumbers.Quantity = batchNumbers1.Quantity;
                                    OINV.Lines.BatchNumbers.Location = batchNumbers1.Location;
                                    OINV.Lines.BatchNumbers.BaseLineNumber = batchNumbers1.BaseLineNumber;
                                }
                            }
                            
                            OINV.Lines.Add();
                        }

                        if (ORDR.TaxExtension.MainUsage != 0)
                        {
                            OINV.TaxExtension.MainUsage = ORDR.TaxExtension.MainUsage;
                        }

                        OINV.TaxExtension.State = ORDR.TaxExtension.State;
                        OINV.TaxExtension.County = ORDR.TaxExtension.County;
                        OINV.TaxExtension.Incoterms = ORDR.TaxExtension.Incoterms;
                        OINV.TaxExtension.Vehicle = ORDR.TaxExtension.Vehicle;
                        OINV.TaxExtension.VehicleState = ORDR.TaxExtension.VehicleState;
                        OINV.TaxExtension.NFRef = ORDR.TaxExtension.NFRef;
                        OINV.TaxExtension.Carrier = ORDR.TaxExtension.Carrier;
                        OINV.TaxExtension.PackQuantity = ORDR.TaxExtension.PackQuantity;
                        OINV.TaxExtension.PackDescription = ORDR.TaxExtension.PackDescription;
                        OINV.TaxExtension.NetWeight = ORDR.TaxExtension.NetWeight;
                        OINV.TaxExtension.GrossWeight = ORDR.TaxExtension.GrossWeight;
                        OINV.TaxExtension.Brand = ORDR.TaxExtension.Brand;
                        OINV.TaxExtension.ShipUnitNo = ORDR.TaxExtension.ShipUnitNo;

                        // ShipTo
                        OINV.AddressExtension.ShipToAddressType = ORDR.AddressExtension.ShipToAddressType;
                        OINV.AddressExtension.ShipToStreet = ORDR.AddressExtension.ShipToStreet;
                        OINV.AddressExtension.ShipToStreetNo = ORDR.AddressExtension.ShipToStreetNo;
                        OINV.AddressExtension.ShipToBuilding = ORDR.AddressExtension.ShipToBuilding;
                        OINV.AddressExtension.ShipToZipCode = ORDR.AddressExtension.ShipToZipCode;
                        OINV.AddressExtension.ShipToBlock = ORDR.AddressExtension.ShipToBlock;
                        OINV.AddressExtension.ShipToCity = ORDR.AddressExtension.ShipToCity;
                        OINV.AddressExtension.ShipToState = ORDR.AddressExtension.ShipToState;
                        OINV.AddressExtension.ShipToCounty = ORDR.AddressExtension.ShipToCounty;
                        OINV.AddressExtension.ShipToCountry = ORDR.AddressExtension.ShipToCountry;
                        OINV.AddressExtension.ShipToGlobalLocationNumber = ORDR.AddressExtension.ShipToGlobalLocationNumber;

                        //BillTo
                        OINV.AddressExtension.BillToAddressType = ORDR.AddressExtension.BillToAddressType;
                        OINV.AddressExtension.BillToStreet = ORDR.AddressExtension.BillToStreet;
                        OINV.AddressExtension.BillToStreetNo = ORDR.AddressExtension.BillToStreetNo;
                        OINV.AddressExtension.BillToBuilding = ORDR.AddressExtension.BillToBuilding;
                        OINV.AddressExtension.BillToZipCode = ORDR.AddressExtension.BillToZipCode;
                        OINV.AddressExtension.BillToBlock = ORDR.AddressExtension.BillToBlock;
                        OINV.AddressExtension.BillToCity = ORDR.AddressExtension.BillToCity;
                        OINV.AddressExtension.BillToState = ORDR.AddressExtension.BillToState;
                        OINV.AddressExtension.BillToCounty = ORDR.AddressExtension.BillToCounty;
                        OINV.AddressExtension.BillToCountry = ORDR.AddressExtension.BillToCountry;
                        OINV.AddressExtension.BillToGlobalLocationNumber = ORDR.AddressExtension.BillToGlobalLocationNumber;

                        int ret1 = OINV.Add();
                        string sErro1 = string.Empty;

                        if (ret1 != 0)
                        {
                            Conexao.diCompany.GetLastError(out ret1, out sErro1);
                            Conexao.sbo_application.StatusBar.SetText($@"Erro para Gerar as Notas [PV00002]: {sErro1}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        }
                        else
                        {
                            Conexao.sbo_application.StatusBar.SetText($@"Nota Fiscal gerada do Pedido {Grid0.DataTable.GetValue("Nº do Doc", k)}"
                                , BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                        }
                        a++;
                    }
                }

                if (a == 0)
                {
                    MessageBox.Show("Nenhum Pedido de Venda selecionado, favor selecionar pelo menos um para gerar a Nota.", "Aviso Gerar a Nota");
                }
            }
            catch (Exception ex)
            {
                Conexao.sbo_application.StatusBar.SetText($@"Erro para Gerar as Notas [PV00003]: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                Conexao.sbo_application.StatusBar.SetText($@"Erro para Gerar as Notas [PV00004]: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                Conexao.sbo_application.StatusBar.SetText($@"Erro ao desmarcar linhas [PV00005]: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                form.Freeze(false);
            }
        }

        private StaticText StaticText0;
    }
}