using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Addon_CopiaPedidoParaNota
{
    [FormAttribute("141", "NotaFiscalEntrada.b1f")]
    class NotaFiscalEntrada : SystemFormBase
    {
        public NotaFiscalEntrada()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("38").Specific));
            this.Matrix0.ComboSelectAfter += new SAPbouiCOM._IMatrixEvents_ComboSelectAfterEventHandler(this.Matrix0_ComboSelectAfter);
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("4").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Matrix Matrix0;
        private EditText EditText0;


        private void Matrix0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                this.UIAPIRawForm.Freeze(true);
                int a = Matrix0.RowCount;

                for (int i = 0; i < a; i++)
                {
                    var ItemCode = ((EditText)(Matrix0.Columns.Item("1").Cells.Item(i + 1).Specific)).Value;
                    var CardCode = EditText0.Value;
                    var endereco = this.UIAPIRawForm.DataSources.DBDataSources.Item("OPCH").GetValue("PayToCode", 0);

                    var ListaPreco = ((SAPbouiCOM.ComboBox)(Matrix0.Columns.Item("U_BDO_GRAU_MEDIO").Cells.Item(i + 1).Specific)).Value;

                    if (!string.IsNullOrEmpty(ItemCode))
                    {
                        SAPbobsCOM.Recordset oRecPrice = (SAPbobsCOM.Recordset)Conexao.diCompany.GetBusinessObject(
                                       SAPbobsCOM.BoObjectTypes.BoRecordset);

                        string queryPrice = $@"SELECT DISTINCT
	                                            T2.""Price""
                                            FROM OPCH T0
	                                            INNER JOIN PCH1 T1 ON T0.""DocEntry"" = T1.""DocEntry""
	                                            INNER JOIN ITM1 T2 ON T1.""ItemCode"" = T2.""ItemCode""
	                                            INNER JOIN OPLN T3 ON T2.""PriceList"" = T3.""ListNum"" 
                                            WHERE
	                                            T1.""ItemCode"" = '{ItemCode}' 
	                                            AND T3.""U_GRAU_UVA"" = {ListaPreco}";

                        oRecPrice.DoQuery(queryPrice);

                        if (oRecPrice.RecordCount > 0)
                        {
                            double Price = Convert.ToDouble(oRecPrice.Fields.Item(0).Value.ToString());

                            string queryMun = $@"SELECT DISTINCT
                                                T0.""County""
                                            FROM CRD1 T0
                                            WHERE
                                                T0.""Address""= '{endereco}'
                                                AND T0.""CardCode"" = '{CardCode}'";

                            SAPbobsCOM.Recordset oRecMunicipio = (SAPbobsCOM.Recordset)Conexao.diCompany.GetBusinessObject(
                                      SAPbobsCOM.BoObjectTypes.BoRecordset);

                            oRecMunicipio.DoQuery(queryMun);

                            if (oRecMunicipio.RecordCount > 0)
                            {
                                string Municipio = oRecMunicipio.Fields.Item(0).Value.ToString();

                                SAPbobsCOM.Recordset oRecMun = (SAPbobsCOM.Recordset)Conexao.diCompany.GetBusinessObject(
                                           SAPbobsCOM.BoObjectTypes.BoRecordset);

                                string sqlMunicipio = $@"SELECT ""U_BDO_VALOR_FRETE"" FROM OCNT WHERE ""AbsId"" = {Municipio}";

                                oRecMun.DoQuery(sqlMunicipio);

                                if (oRecMun.RecordCount > 0)
                                {
                                    double valorAtual = Price + Convert.ToDouble(oRecMun.Fields.Item(0).Value.ToString());

                                    ((EditText)(Matrix0.Columns.Item("14").Cells.Item(i + 1).Specific)).Value = valorAtual.ToString();
                                }
                            }
                            else
                            {
                                ((EditText)(Matrix0.Columns.Item("14").Cells.Item(i + 1).Specific)).Value = "";
                            }
                        }
                        else
                        {
                            ((EditText)(Matrix0.Columns.Item("14").Cells.Item(i + 1).Specific)).Value = "";
                        }
                    }
                }
                this.UIAPIRawForm.Freeze(false);
            }
            catch (Exception ex)
            {
                this.UIAPIRawForm.Freeze(false);
                MessageBox.Show("ANFE0001: " + ex.Message, "Addon Nota fiscal Entrada - Grau Uva");
            }

        }

        private void OnCustomInitialize()
        {

        }


    }
}
