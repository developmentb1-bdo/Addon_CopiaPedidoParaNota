using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Addon_CopiaPedidoParaNota
{
    class CreateFields
    {
        public void EventHandlerStart()
        {
            Conexao oCon = new Conexao();
            try
            {
                Application.SBO_Application.StatusBar.SetText("Iniciando criação de Campos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                
                string[,] vvStatus = new string[13, 2];
                vvStatus[0, 0] = "0";
                vvStatus[0, 1] = "0";
                vvStatus[1, 0] = "10";
                vvStatus[1, 1] = "10";
                vvStatus[2, 0] = "11";
                vvStatus[2, 1] = "11";
                vvStatus[3, 0] = "12";
                vvStatus[3, 1] = "12";
                vvStatus[4, 0] = "13";
                vvStatus[4, 1] = "13";
                vvStatus[5, 0] = "14";
                vvStatus[5, 1] = "14";
                vvStatus[6, 0] = "15";
                vvStatus[6, 1] = "15";
                vvStatus[7, 0] = "16";
                vvStatus[7, 1] = "16";
                vvStatus[8, 0] = "17";
                vvStatus[8, 1] = "17";
                vvStatus[9, 0] = "18";
                vvStatus[9, 1] = "18";
                vvStatus[10, 0] = "19";
                vvStatus[10, 1] = "19";
                vvStatus[11, 0] = "20";
                vvStatus[11, 1] = "20";
                vvStatus[12, 0] = "21";
                vvStatus[12, 1] = "21";

                Application.SBO_Application.StatusBar.SetText("Criando Campo Valor Frete.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                oCon.AddUserField("OCNT", "BDO_VALOR_FRETE", "Valor Frete", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, null, null, null);

                Application.SBO_Application.StatusBar.SetText("Criando Campo GRAU da uva.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                oCon.AddUserField("PCH1", "BDO_GRAU_MEDIO", "GRAU da uva para o arquivo SISDEVIN", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 2, vvStatus, "0", null);

                Application.SBO_Application.StatusBar.SetText("Campos criado com sucesso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception)
            {
                Environment.Exit(0);
            }
        }
    }
}
