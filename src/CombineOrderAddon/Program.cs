using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;

namespace CombineOrderAddon
{
    class Program
    {
        static SAPbobsCOM.Company oCom;
        static SAPbobsCOM.Recordset oRS;
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    oApp = new Application(args[0]);
                }
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;

                oCom = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                oRS = (SAPbobsCOM.Recordset)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.FormTypeEx == "139" && pVal.BeforeAction == false && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                SAPbouiCOM.Item oButtonPurchase = oForm.Items.Add("action", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                SAPbouiCOM.Item oTempItem = oForm.Items.Item("30");
                SAPbouiCOM.Button oPostButton = (SAPbouiCOM.Button)oButtonPurchase.Specific;

                oPostButton.Caption = "Заказ бирлаштириш";
                oButtonPurchase.Left = oTempItem.Left;
                oButtonPurchase.Top = oTempItem.Top + 60;
                oButtonPurchase.Width = 130;
                oButtonPurchase.Height = oTempItem.Height + 5;
                oButtonPurchase.AffectsFormMode = false;
            }

            if (pVal.FormTypeEx == "139" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "action" && pVal.BeforeAction == false)
            {
                try
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

                    Application.SBO_Application.StatusBar.SetSystemMessage("Начался процесс генерации ...",
                        SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                    SAPbouiCOM.EditText importEnt = (SAPbouiCOM.EditText)oForm.Items.Item("134").Specific;

                    oRS.DoQuery($"SELECT T1.\"LineNum\", T0.\"DocNum\", T0.\"DocEntry\", T0.\"CardCode\", T0.\"CardName\", T0.\"CntctCode\", T0.\"DocDate\", T0.\"DocDueDate\", " +
                        $"T0.\"TaxDate\", T0.\"ImportEnt\", T1.\"ItemCode\", T1.\"Dscription\", T1.\"Quantity\", T1.\"Price\", T1.\"DiscPrcnt\", T1.\"WhsCode\" " +
                        $"FROM ORDR AS T0 INNER JOIN RDR1 AS T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" WHERE T0.\"ImportEnt\" = '{importEnt.Value}'");


                    SAPbobsCOM.Documents invoice = (SAPbobsCOM.Documents)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                    invoice.CardCode = oRS.Fields.Item("CardCode").Value.ToString();
                    invoice.CardName = oRS.Fields.Item("CardName").Value.ToString();
                    invoice.ContactPersonCode = int.Parse(oRS.Fields.Item("CntctCode").Value.ToString());
                    invoice.DocDate = DateTime.Parse(oRS.Fields.Item("DocDate").Value.ToString());
                    invoice.DocDueDate = DateTime.Parse(oRS.Fields.Item("DocDueDate").Value.ToString());
                    invoice.TaxDate = DateTime.Parse(oRS.Fields.Item("TaxDate").Value.ToString());

                    for (int i = 1; i <= oRS.RecordCount; i++)
                    {
                        var item = invoice.Lines;
                        item.ItemCode = oRS.Fields.Item("ItemCode").Value.ToString();
                        item.ItemDescription = oRS.Fields.Item("Dscription").Value.ToString();
                        item.Quantity = double.Parse(oRS.Fields.Item("Quantity").Value.ToString());
                        item.UnitPrice = double.Parse(oRS.Fields.Item("Price").Value.ToString());
                        item.DiscountPercent = double.Parse(oRS.Fields.Item("DiscPrcnt").Value.ToString());
                        item.WarehouseCode = oRS.Fields.Item("WhsCode").Value.ToString();

                        item.BaseLine = int.Parse(oRS.Fields.Item("LineNum").Value.ToString());
                        item.BaseEntry = int.Parse(oRS.Fields.Item("DocEntry").Value.ToString());
                        item.BaseType = 17;

                        // Console.WriteLine($"{item.ItemCode}    {item.ItemDescription}   {item.Quantity}   {item.WarehouseCode}    {item.BaseLine}   {item.BaseEntry}");
                        item.Add();
                        oRS.MoveNext();
                    }

                    int status = invoice.Add();

                    if (status != 0)
                    {
                        int errorCode = oCom.GetLastErrorCode();
                        string error = oCom.GetLastErrorDescription();
                        Application.SBO_Application.StatusBar.SetSystemMessage("Ошибка при генерации", 
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        Console.WriteLine($"{errorCode} {error}");
                    }
                    else
                    {
                        Application.SBO_Application.StatusBar.SetSystemMessage($"Генерации завершена",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }

                }
                catch (Exception exception)
                {
                    Console.WriteLine(exception.Message);
                }
            }
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }
    }
}
