using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Approved_BPCreation
{
    class ClsABPGroupControl : Base
    {
        //Variable Declaration
        bool flag = false;
        string groupCode = "";
        public ClsABPGroupControl() : base()
        { }

        ~ClsABPGroupControl() { }


        #region ITEM EVENT
        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {

            #region BEFORE ACTION ACTION

            if (pVal.Before_Action == true)
            {
                switch (pVal.EventType)
                {
                    #region ITEM PRESS EVENT
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && (_Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || _Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
                        {
                            flag = true;
                        }
                        break;
                        #endregion
                }
            }

            #endregion

            #region After ACTION ACTION

            else if (pVal.Before_Action == false)
            {

                switch (pVal.EventType)
                {
                    #region ITEM PRESS EVENT
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && pVal.ActionSuccess == true && flag == true)
                        {
                            SyncAllOCRGToUDO();
                        }
                        break;
                        #endregion
                }
            }

            #endregion

        }
        #endregion
        public void FormDefault()
        {
            try
            {

            }
            catch(Exception ex)
            {

            }

        }

        public void SyncAllOCRGToUDO()
        {
            SAPbobsCOM.Company oCompany = Utilities.Application.Company;

            SAPbobsCOM.Recordset rsOCRG = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            SAPbobsCOM.Recordset rsCheck = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // Fetch all BP Groups
            rsOCRG.DoQuery(@"SELECT ""GroupCode"", ""GroupName"", ""GroupType"" FROM OCRG");

            SAPbobsCOM.GeneralService oService =(SAPbobsCOM.GeneralService)oCompany.GetCompanyService().GetGeneralService("ABPGROUPC"); // UDO Object Name

            while (!rsOCRG.EoF)
            {
                string groupCode = rsOCRG.Fields.Item("GroupCode").Value.ToString();
                string groupName = rsOCRG.Fields.Item("GroupName").Value.ToString();
                string groupType = rsOCRG.Fields.Item("GroupType").Value.ToString(); 

                // Check if record exists in UDO
                rsCheck.DoQuery($@"
            SELECT ""Code""
            FROM ""@ABPGROUPC""
            WHERE ""Code"" = '{groupCode}'");

                if (rsCheck.RecordCount > 0)
                {
                    //UPDATE
                    SAPbobsCOM.GeneralDataParams oParams = (SAPbobsCOM.GeneralDataParams)oService.GetDataInterface( SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);

                    oParams.SetProperty("Code", groupCode);

                    SAPbobsCOM.GeneralData oData = oService.GetByParams(oParams);
                    oData.SetProperty("Name", groupName);

                    oService.Update(oData);
                }
                else
                {
                    //ADD
                    SAPbobsCOM.GeneralData oData = (SAPbobsCOM.GeneralData)oService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                    oData.SetProperty("Code", groupCode);
                    oData.SetProperty("Name", groupName);
                    oData.SetProperty("U_BPApplicable", "No");
                    oData.SetProperty("U_GroupType", groupType);


                    oService.Add(oData);
                }

                rsOCRG.MoveNext();
            }

            Utilities.Application.SBO_Application.StatusBar.SetText( "Vendor Groups successfully synced.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        }



    }
}
