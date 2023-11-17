using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Addon_CopiaPedidoParaNota
{
    public class Conexao
    {
        public static SAPbobsCOM.Company diCompany;

        public static SAPbouiCOM.Application sbo_application = SAPbouiCOM.Framework.Application.SBO_Application;

        public Conexao()
        {
            try
            {
                this.SetApplicationDI();
            }
            catch
            {
                throw;
            }
        }
        public void SetApplicationDI()
        {
            try
            {
                int num;
                diCompany = (SAPbobsCOM.Company)SAPbouiCOM.Framework.Application.SBO_Application.Company.GetDICompany();
                string errMsg = "";

                diCompany.GetLastError(out num, out errMsg);
                if (num != 0)
                {
                    throw new Exception(errMsg);
                }
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }
        public static object ExecuteSqlScalar(string query)
        {
            object obj3;
            try
            {
                object obj2 = null;
                SAPbobsCOM.Recordset businessObject = (SAPbobsCOM.Recordset)diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                businessObject.DoQuery(query);
                if (!businessObject.EoF)
                {
                    obj2 = businessObject.Fields.Item(0).Value;
                }
                Marshal.ReleaseComObject(businessObject);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                obj3 = obj2;
            }
            catch (Exception)
            {
                throw;
            }
            return obj3;
        }
        public void AddUserTable(string NomeTB, string Desc, SAPbobsCOM.BoUTBTableType oTableType)
        {
            int lErrCode;
            string sErrMsg = "";

            SAPbobsCOM.UserTablesMD oUserTable = (SAPbobsCOM.UserTablesMD)diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
            Marshal.ReleaseComObject(oUserTable);
            oUserTable = null;

            oUserTable = (SAPbobsCOM.UserTablesMD)diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

            try
            {
                if (!oUserTable.GetByKey(NomeTB))
                {
                    oUserTable.TableName = NomeTB.Replace("@", "").Replace("[", "").Replace("]", "").Trim();
                    oUserTable.TableDescription = Desc;
                    oUserTable.TableType = oTableType;

                    try
                    {
                        if (oUserTable.Add() != 0)
                        {
                            diCompany.GetLastError(out lErrCode, out sErrMsg);
                            sbo_application.SetStatusBarMessage($@"Erro ao criar tabela - {NomeTB} - {sErrMsg}", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                            throw new Exception("Erro: " + sErrMsg);
                        }
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                Marshal.ReleaseComObject(oUserTable);
                oUserTable = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }
        public void AddUserField(string NomeTabela, string NomeCampo, string DescCampo, SAPbobsCOM.BoFieldTypes Tipo, SAPbobsCOM.BoFldSubTypes SubTipo,
                                    Int16 Tamanho, string[,] valoresValidos, string valorDefault, string linkedTable)
        {
            int lErrCode;
            string sErrMsg = "";

            string strSql = string.Format(@"select COUNT(*)  
                                                from CUFD 
                                                where ""TableID"" = '{0}' 
                                                    and ""AliasID"" = '{1}'", NomeTabela, NomeCampo);
            //0 - Campo Não exite
            //1 - Campos Existe
            int resultado = (int)ExecuteSqlScalar(strSql);
            if (resultado == 0)
            {
                try
                {
                    //string sSquery = "SELECT ""[name]"" FROM syscolumns WHERE ""[name]"" = 'U_" + NomeCampo + " ' and id = (SELECT id FROM sysobjects WHERE type = 'U'AND [NAME] = '" + NomeTabela.Replace("[", "").Replace("]", "") + "')";
                    //object oResult = B1Connections.ExecuteSqlScalar(sSquery);
                    //if (oResult != null) return;

                    SAPbobsCOM.UserFieldsMD oUserField;
                    oUserField = (SAPbobsCOM.UserFieldsMD)Conexao.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                    oUserField.TableName = NomeTabela.Replace("@", "").Replace("[", "").Replace("]", "").Trim();
                    oUserField.Name = NomeCampo;
                    oUserField.Description = DescCampo;
                    oUserField.Type = Tipo;
                    oUserField.SubType = SubTipo;
                    oUserField.DefaultValue = valorDefault;
                    if (!string.IsNullOrEmpty(linkedTable)) oUserField.LinkedTable = linkedTable;

                    //adicionar valores válidos
                    if (valoresValidos != null)
                    {
                        Int32 qtd = valoresValidos.GetLength(0);
                        if (qtd > 0)
                        {
                            for (int i = 0; i < qtd; i++)
                            {
                                oUserField.ValidValues.Value = valoresValidos[i, 0];
                                oUserField.ValidValues.Description = valoresValidos[i, 1];
                                oUserField.ValidValues.Add();
                            }
                        }
                    }

                    if (Tamanho != 0)
                        oUserField.EditSize = Tamanho;

                    try
                    {
                        oUserField.Add();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField);
                        oUserField = null;
                        Conexao.diCompany.GetLastError(out lErrCode, out sErrMsg);
                        if (lErrCode != 0)
                        {
                            Conexao.sbo_application.StatusBar.SetText($@"Erro ao criar campo - {NomeCampo} - {sErrMsg}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            throw new Exception(sErrMsg);
                        }
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                    oUserField = null;
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
        }
        public bool ExisteTB(string TBName)
        {
            SAPbobsCOM.UserTablesMD oUserTable;
            oUserTable = (SAPbobsCOM.UserTablesMD)Conexao.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
            //UserTablesMD oUserTable = new UserTablesMD(ref oDiCompany);            
            bool ret = oUserTable.GetByKey(TBName);
            int errCode; string errMsg;
            Conexao.diCompany.GetLastError(out errCode, out errMsg);

            TBName = null;
            errMsg = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            return (ret);
        }
        public void AddUDO(string sUDO, string sTable, string sDescricaoUDO, SAPbobsCOM.BoUDOObjType oBoUDOObjType, string[] childTableName, string[] childObjectName)
        {
            int lRetCode = 0;
            int iTabelasFilhas = 0;
            string sErrMsg = "";
            bool bUpdate = false;
            bool bExisteTabelaFilha = false;

            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;

            oUserObjectMD = (SAPbobsCOM.UserObjectsMD)Conexao.diCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

            System.Data.DataTable tb = new System.Data.DataTable();

            try
            {
                if (oUserObjectMD.GetByKey(sUDO))
                {
                    return;
                }

                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.Code = sUDO;
                oUserObjectMD.Name = sDescricaoUDO;
                oUserObjectMD.ObjectType = oBoUDOObjType;
                oUserObjectMD.TableName = sTable;

                //Adicionar tabelas filhas
                if (childObjectName != null)
                {
                    for (int x = 0; x < childObjectName.Length; x++)
                    {

                        iTabelasFilhas = oUserObjectMD.ChildTables.Count;
                        bExisteTabelaFilha = false;
                        for (int y = 0; y < iTabelasFilhas; y++)
                        {
                            oUserObjectMD.ChildTables.SetCurrentLine(y);
                            if (oUserObjectMD.ChildTables.TableName == childTableName[x])
                            {
                                bExisteTabelaFilha = true;
                                break;
                            }
                        }

                        if (bExisteTabelaFilha == false)
                        {
                            if (x > 0) oUserObjectMD.ChildTables.Add();
                            if (childObjectName[x] != "" && childTableName[x] != "")
                            {
                                oUserObjectMD.ChildTables.TableName = childTableName[x];
                                oUserObjectMD.ChildTables.ObjectName = childObjectName[x];
                            }
                        }

                    }
                }

                if (bUpdate)
                    lRetCode = oUserObjectMD.Update();
                else
                    lRetCode = oUserObjectMD.Add();

                // check for errors in the process
                if (lRetCode != 0)
                {
                    Conexao.diCompany.GetLastError(out lRetCode, out sErrMsg);
                }

            }
            catch (Exception e)
            { System.Windows.Forms.MessageBox.Show(e.ToString()); }


            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
            oUserObjectMD = null;
            tb.Dispose();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        public static void CentralizarForm(SAPbouiCOM.Form oForm)
        {
            oForm.Left = (SAPbouiCOM.Framework.Application.SBO_Application.Desktop.Width - oForm.Width) / 2;
            oForm.Top = ((SAPbouiCOM.Framework.Application.SBO_Application.Desktop.Height - oForm.Height) / 2) - 100;
        }
    }
}
