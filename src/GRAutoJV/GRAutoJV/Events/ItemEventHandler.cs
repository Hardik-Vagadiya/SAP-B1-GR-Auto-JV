using GRAutoJV.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GRAutoJV.Models;
namespace GRAutoJV.Events
{
    public class ItemEventHandler
    {
        public static void ChPoValidate(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.FormTypeEx == "721" && pVal.ItemUID == "1" && pVal.BeforeAction &&
                pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    SAPbouiCOM.Form oForm = ConnectionManager.oApp.Forms.Item(pVal.FormUID);
                    ConnectionManager.oGlobalVars = new GlobalVars();
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        ConnectionManager.oGlobalVars.jvValue = 0;
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("13").Specific;
                        SAPbobsCOM.Recordset oRes = (SAPbobsCOM.Recordset)ConnectionManager.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        int rowCount = oMatrix.RowCount;
                        SAPbouiCOM.Column oColPrice = oMatrix.Columns.Item("10");
                        oColPrice.Editable = true;
                        for (int i = 1; i < rowCount; i++)
                        {
                            SAPbouiCOM.EditText JwPo = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_JwPoNum").Cells.Item(i).Specific;
                            SAPbouiCOM.EditText ChNo = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_InvntryTrans").Cells.Item(i).Specific;
                            SAPbouiCOM.EditText rQty = (SAPbouiCOM.EditText)oMatrix.Columns.Item("9").Cells.Item(i).Specific;
                            SAPbouiCOM.EditText rNetWt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_NetWt").Cells.Item(i).Specific;
                            SAPbouiCOM.ComboBox recType = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("U_RecType").Cells.Item(i).Specific;
                            SAPbouiCOM.EditText RecCode = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific;
                            SAPbouiCOM.EditText RecDesc = (SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(i).Specific;

                            if (string.IsNullOrEmpty(ChNo.Value.ToString()))
                            {
                                ConnectionManager.oApp.StatusBar.SetText("Inventory Transfer cannot be blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                            if (string.IsNullOrEmpty(JwPo.Value.ToString()) && !(RecDesc.Value.ToLower().Contains("scrap"))
                                && !( recType.Value.ToString() == "Rew" || recType.Value.ToString() == "WoP"))
                            {
                                ConnectionManager.oApp.StatusBar.SetText("JW PO Number cannot be blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }

                            double grQty, grNetWt;
                            bool success = double.TryParse(rQty.Value, out grQty);
                            bool success1 = double.TryParse(rNetWt.Value, out grNetWt);
                            if (!(success && success1))
                            {
                                ConnectionManager.oApp.StatusBar.SetText("Invalid quantity/Net weight!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }

                            if (grQty <= 0 || grNetWt <= 0)
                            {
                                ConnectionManager.oApp.StatusBar.SetText("Quantity/Net wt should be greater than zero!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }

                            string queryIT = "SELECT \"ItemCode\", \"Dscription\", \"StockPrice\", \"Quantity\" FROM" +
                                " WTR1 WHERE concat(\"DocEntry\",\"LineNum\") = " + Convert.ToInt32(ChNo.Value);

                            string queryPO = "SELECT MIN(\"Price\") as \"Price\" FROM" +
                                " POR1 WHERE concat(\"DocEntry\", \"LineNum\") = " + Convert.ToInt32(JwPo.Value) 
                                + " AND \"U_Item\" = " + "'" + RecCode.Value.ToString() + "'"
                                + " AND NOT(\"ItemCode\" LIKE 'SFG%' AND \"unitMsr\" IN ('KG', 'Kg', 'Kgs', 'Kilogram'));";

                            oRes.DoQuery(queryIT);
                            if (oRes.RecordCount <= 0)
                            {
                                BubbleEvent = false;
                                ConnectionManager.oApp.StatusBar.SetText($"Row {i}: No challan record found!.",
                                SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                return;
                            }
                            double qty = Convert.ToDouble(oRes.Fields.Item("Quantity").Value.ToString());
                            double price = Convert.ToDouble(oRes.Fields.Item("StockPrice").Value.ToString());
                            
                            string IssCode = oRes.Fields.Item("ItemCode").Value.ToString();
                            
                            string IssDesc = oRes.Fields.Item("Dscription").Value.ToString();
                            double PoRate = 0;

                            if (!(string.IsNullOrEmpty(JwPo.Value.ToString())))
                            {
                                oRes.DoQuery(queryPO);
                                if (!oRes.EoF && oRes.Fields.Item("Price").Value != null)
                                {
                                    PoRate = Convert.ToDouble(oRes.Fields.Item("Price").Value);
                                }
                            }
                            //double PoRate = Convert.ToDouble(oRes.Fields.Item("Price").Value.ToString());

                            if (qty == 0 || price == 0)
                            {
                                BubbleEvent = false;
                                ConnectionManager.oApp.StatusBar.SetText($"Row {i}: Challan qty/price cannot be null!. {qty}, {price}",
                                SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                return;
                            }

                            double InPrice = 0;

                            if (IssCode.Substring(0, 3) == "RMG" && !(IssDesc.ToLower().Contains("scrap"))
                                && RecCode.Value.ToString().Substring(0, 3) == "SFG")
                            {
                                InPrice = ((grNetWt * price) / grQty) + PoRate;
                                ConnectionManager.oGlobalVars.jvValue += Convert.ToDouble(rQty.Value) * PoRate;
                            }
                            else if (IssCode.Substring(0, 3) == "RMG" && !(IssDesc.ToLower().Contains("scrap"))
                                && RecCode.Value.ToString().Substring(0, 3) == "RMG" && RecDesc.Value.ToString().ToLower().Contains("scrap"))
                            {                             
                                SAPbobsCOM.Recordset itmRes = (SAPbobsCOM.Recordset)ConnectionManager.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                string query = "SELECT \"LstEvlPric\" FROM OITM WHERE \"ItemCode\" = " + "'" + IssCode.ToString() + "'";
                                itmRes.DoQuery(query);
                                if (!itmRes.EoF)
                                {
                                    InPrice = Convert.ToDouble(itmRes.Fields.Item("LstEvlPric").Value);
                                }
                                else
                                {
                                    InPrice = 0;
                                }
                            }
                            else if (IssCode.Substring(0, 3) == "RMG" && !(IssDesc.ToLower().Contains("scrap"))
                                && RecCode.Value.ToString().Substring(0, 3) == "RMG" && !(RecDesc.Value.ToString().ToLower().Contains("scrap")))
                            {
                                InPrice = price;
                            }
                            else if (IssCode.Substring(0, 3) == "RMG" && (IssDesc.ToLower().Contains("scrap"))
                                && RecCode.Value.ToString().Substring(0, 3) == "RMG" && !(RecDesc.Value.ToString().ToLower().Contains("scrap")))
                            {
                                InPrice = price + PoRate;
                                ConnectionManager.oGlobalVars.jvValue += PoRate * Convert.ToDouble(rQty.Value);
                            }
                            else if (IssCode.Substring(0, 3) == "RMG" && (IssDesc.ToLower().Contains("scrap"))
                                && RecCode.Value.ToString().Substring(0, 3) == "RMG" && (RecDesc.Value.ToString().ToLower().Contains("scrap")))
                            {                                
                                InPrice = price;
                            }
                            else if (IssCode.Substring(0, 3) == "SFG" && RecCode.Value.ToString().Substring(0, 3) == "SFG")
                            {                                
                                InPrice = price + PoRate;
                                ConnectionManager.oGlobalVars.jvValue += PoRate * Convert.ToDouble(rQty.Value);
                            }
                            else if (IssCode.Substring(0, 3) == "SFG" && RecCode.Value.ToString().Substring(0, 3) == "RMG"
                                && (RecDesc.Value.ToString().ToLower().Contains("scrap")))
                            {
                                SAPbobsCOM.Recordset itmRes = (SAPbobsCOM.Recordset)ConnectionManager.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                string query = "SELECT \"LstEvlPric\" FROM OITM WHERE \"ItemCode\" = " + "'" + IssCode.ToString() + "'";
                                itmRes.DoQuery(query);
                                if (!itmRes.EoF)
                                {
                                    InPrice = Convert.ToDouble(itmRes.Fields.Item("LstEvlPric").Value);
                                }
                                else
                                {
                                    InPrice = 0;
                                }
                            }
                            else
                            {
                                InPrice = 0;
                            }
                            SAPbouiCOM.EditText uPrice = (SAPbouiCOM.EditText)oMatrix.Columns.Item("10").Cells.Item(i).Specific;
                            uPrice.Value = InPrice.ToString();
                        }
                        oForm.Items.Item("11").Click();
                        oColPrice.Editable = false;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
    }
}
