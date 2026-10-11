using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sap.Data.Hana;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Xml;
using System.Collections;
using System.Threading;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Globalization;
using Sap.Data.Hana;
using System.Data.SqlClient;
using System.Data;

namespace Approved_BPCreation
{
    class ClsApprovalTemplate : Base
    {
        #region VARIABLE DECLARATION

        SAPbouiCOM.Item _ExistItem, _NewItem, _Items;
        private bool flag;
        private bool errorFlag;
        public SAPbouiCOM.Matrix _Matrix;
        SAPbouiCOM.StaticText _StaticText;
        SAPbouiCOM.Button _Button;
        string _DocEntry = "";
        string _PRNum = string.Empty, _PRName = string.Empty, _RevNum = string.Empty, _Warehouse = string.Empty, _RevRmrks = string.Empty;
        double _Qty = 0;
        public string ErrMsg = string.Empty;
        public int ErrCode = 0;
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        private static HanaConnection _conn;
        private static HanaDataAdapter _da;
        //private static DataTable _dataTable;
        private static HanaParameter _parm1;
        private static HanaParameter _parm2;
        private static int _errors = 0;
        string strcon, _Sql;
        private SAPbouiCOM.EditText _EditText;
        private SAPbouiCOM.LinkedButton _Linked;
        private SAPbouiCOM.ComboBox _ComboBoxM;
        private SAPbouiCOM.ComboBox _ComboBoxC, _ComboBoxCRow;
        private SAPbouiCOM.ComboBox _ComboBoxB;
        private SAPbouiCOM.ComboBox _ComboBoxL;
        private SAPbouiCOM.ComboBox _ComboBoxH;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource = null, _ChildDBDataSource = null;
        private SAPbouiCOM.ChooseFromListCollection _CFLS;
        private SAPbouiCOM.ChooseFromList _CFL;
        private SAPbouiCOM.ChooseFromListCreationParams _CFLC;
        private SAPbouiCOM.Conditions _Cons;
        private SAPbouiCOM.Condition _Con;
        private int RowID { get; set; }
        private string _Code { get; set; }
        private string _LineId { get; set; }
        private string _ActivityType = string.Empty;
        string _ItemCode = string.Empty, _ItemName = string.Empty;
        private SAPbobsCOM.Recordset _Rs, _RS, _rs, _RSD, _RSBom, _RSProd, _RSRoute, _RSUsr, _RSMgItm, _RSBomP;
        SAPbouiCOM.ChooseFromListEvent _SysCFLEvent = null;
        SAPbouiCOM.DBDataSource _DBData = null;
        private System.Collections.Generic.List<string> ListOfBlockedID;
        internal string _ServiceID { get; set; }
        internal string _ActivityTypes { get; set; }
        public bool chkStatus = true;
        static int StageID = 0;
        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsApprovalTemplate()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsApprovalTemplate()
        {
        }
        #endregion

        #region ItemEvent
        public override void Item_Event(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
        {
            #region Before Action
            if (pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {

                }
            }
            #endregion

            #region After Action
            else
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                        if (pVal.ItemUID == "btnApprove")
                        {
                            ProcessSelections("Approved");

                        }
                        if (pVal.ItemUID == "btnReject")
                        {
                            ProcessSelections("Draft");

                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:

                        //var matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("matrix").Specific;
                        //SAPbouiCOM.Column column = matrix.Columns.Item("colDoc");
                        //SAPbouiCOM.LinkedButton linkedButton = (SAPbouiCOM.LinkedButton)column.ExtendedObject;
                        //linkedButton.LinkedObject = (BoLinkedObject)SAPbouiCOM.BoUDOObjectType.udo_Document;
                        //linkedButton.LinkedObjectType = "ProductionP"; 

                        if (pVal.EventType == BoEventTypes.et_MATRIX_LINK_PRESSED)
                        {
                            if (pVal.ColUID == "colDoc")
                            {
                                SAPbouiCOM.Form form = Utilities.Application.SBO_Application.Forms.Item(pVal.FormUID);
                                SAPbouiCOM.Matrix matrix = (SAPbouiCOM.Matrix)form.Items.Item("matrix").Specific;

                                string docEntry = ((SAPbouiCOM.EditText)matrix.Columns.Item("colDoc").Cells.Item(pVal.Row).Specific).Value;

                                if (!string.IsNullOrEmpty(docEntry))
                                {
                                    try
                                    {
                                        Utilities.Application.SBO_Application.OpenForm(BoFormObjectEnum.fo_UserDefinedObject, "ProductionP", docEntry);
                                        form.Close();
                                    }
                                    catch (Exception ex)
                                    {
                                        Utilities.Application.SBO_Application.StatusBar.SetText($"Error opening UDO form: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    }
                                }
                            }
                        }


                        break;
                }
            }

            #endregion
        }
        #endregion

        #region FORM DEFAULT
        internal void FormDefault()
        {
            string userCode = Utilities.Application.SBO_Application.Company.UserName;
            Utilities.ExecuteSQL(ref _RSUsr, $"SELECT \"U_BPAuth\" FROM \"OUSR\" WHERE \"USER_CODE\" = '{userCode.Replace("'", "''")}'");


            // Step 4: Check authorization
            if (!_RSUsr.EoF && _RSUsr.Fields.Item("U_BPAuth").Value.ToString() == "Yes")
            {
                LoadApprovalData(userCode);

            }
            else
            {
                // Show error message
                Utilities.Application.SBO_Application.MessageBox("You are not an approver.", 1, "OK", "", "");
            }

            // Cleanup
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_RSUsr);
            _RSUsr = null;

        }
        #endregion

        #region LoadApprovalData
        private void LoadApprovalData(string userCode)
        {
            var matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("matrix").Specific;

            // Step 1: Prepare DataTable (with safe check)
            SAPbouiCOM.DataTable dt;

            try
            {
                dt = _Form.DataSources.DataTables.Item("dtDisplay");
                dt.Clear();  // reuse if already exists
            }
            catch
            {
                // create if doesn't exist
                dt = _Form.DataSources.DataTables.Add("dtDisplay");
                dt.Columns.Add("U_Selected", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 1);
                dt.Columns.Add("U_BPCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 30);
                dt.Columns.Add("U_BPType", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_Status", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
                //dt.Columns.Add("U_Remarks", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254);
                dt.Columns.Add("U_CretedBy", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_ARemarks", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254);
            }

            // Step 2: Bind Matrix Columns to DataTable
            matrix.Columns.Item("colBpCode").DataBind.Bind("dtDisplay", "U_BPCode");
            matrix.Columns.Item("colBpType").DataBind.Bind("dtDisplay", "U_BPType");
            matrix.Columns.Item("colStatus").DataBind.Bind("dtDisplay", "U_Status");
            //matrix.Columns.Item("colRmrks").DataBind.Bind("dtDisplay", "U_Remarks");
            matrix.Columns.Item("colCretBy").DataBind.Bind("dtDisplay", "U_CretedBy");
            matrix.Columns.Item("colAPRmrks").DataBind.Bind("dtDisplay", "U_ARemarks");
            string sql = $@"

                            select ""Code"", ""U_BPType"", ""U_Status"", ""U_CreatedBy"", ""U_Remarks"" from ""@ABPMASTERH"" where ""U_Status"" = 'Awaiting Approval'";
            var rs = (SAPbobsCOM.Recordset)Utilities.Application.Company
                         .GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(sql);

            // Step 4: Load Data into DataTable
            while (!rs.EoF)
            {
                int row = dt.Rows.Count;
                dt.Rows.Add();
                dt.SetValue("U_BPCode", row, rs.Fields.Item("Code").Value.ToString());
                dt.SetValue("U_BPType", row, rs.Fields.Item("U_BPType").Value.ToString());
                dt.SetValue("U_Status", row, rs.Fields.Item("U_Status").Value.ToString());
                //dt.SetValue("U_Remarks", row, rs.Fields.Item("U_Remarks").Value.ToString());
                dt.SetValue("U_CretedBy", row, rs.Fields.Item("U_CreatedBy").Value.ToString());
                dt.SetValue("U_ARemarks", row, rs.Fields.Item("U_Remarks").Value.ToString());

                rs.MoveNext();
            }

            // Step 5: Load Matrix from DataTable
            matrix.LoadFromDataSource();
        }
        #endregion

        #region ProcessSelections
        private void ProcessSelections(string newStatus)
        {
            var matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("matrix").Specific;
            var dt = _Form.DataSources.DataTables.Item("dtDisplay");

            // 1) Get the CompanyService and GeneralService for your UDO
            var companyService = Utilities.Application.Company.GetCompanyService();
            var svc = (SAPbobsCOM.GeneralService)companyService.GetGeneralService("APPROVEDBPC");

            // 2) Retrieve the correct GeneralDataParams interface
            SAPbobsCOM.GeneralDataParams generalParams;
            try
            {
                generalParams = (SAPbobsCOM.GeneralDataParams)svc.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
            }
            catch
            {
                generalParams = (SAPbobsCOM.GeneralDataParams)svc.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralCollectionParams);
            }

            // 3) Iterate selected rows and update UDO
            for (int i = 1; i <= matrix.RowCount; i++)
            {
                var chkColumn = matrix.Columns.Item("chkSelect");
                var cell = chkColumn.Cells.Item(i).Specific;

                bool isSelected = false;
                if (chkColumn.Type == SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
                {
                    var chk = (SAPbouiCOM.CheckBox)cell;
                    isSelected = chk.Checked;
                }
                else if (chkColumn.Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                {
                    var combo = (SAPbouiCOM.ComboBox)cell;
                    isSelected = combo.Selected.Value == "Y";
                }

                if (isSelected)
                {

                    var matrixCol = matrix.Columns.Item("colBpCode");
                    var cellSpec = matrixCol.Cells.Item(i).Specific;


                    var txt = (SAPbouiCOM.EditText)cellSpec;
                    //string docEntryStr = txt.Value?.Trim();
                    string BpCode = txt.Value?.Trim();

                    var matrixColR = matrix.Columns.Item("colAPRmrks");
                    var cellSpecR = matrixColR.Cells.Item(i).Specific;


                    var txtR = (SAPbouiCOM.EditText)cellSpecR;
                    //string docEntryStr = txt.Value?.Trim();
                    string Remarks= txtR.Value?.Trim();


                    //if (!int.TryParse(docEntryStr, out int BpCode))
                    //{
                    //    Utilities.Application.SBO_Application.MessageBox($"Row {i}: Invalid BPCode value '{docEntryStr}'");
                    //    continue;
                    //}
                    if (string.IsNullOrEmpty(BpCode))
                    {
                        Utilities.Application.SBO_Application.MessageBox($"Row {i}: BPCode is empty or invalid.");
                        continue;
                    }
                    generalParams.SetProperty("Code", BpCode);
                    var data = svc.GetByParams(generalParams);
                    data.SetProperty("U_Status", newStatus);
                    data.SetProperty("U_Remarks", Remarks);
                    if (newStatus == "Draft")
                    {
                        data.SetProperty("U_IsRejected", "Yes");
                    }
                    //var remarksCell = (SAPbouiCOM.EditText)matrix.Columns.Item("colAPRmrks").Cells.Item(i).Specific;
                    //string remarksValue = remarksCell.Value?.Trim() ?? "";
                    //data.SetProperty("U_ARemarks", remarksValue);
                    svc.Update(data);
                    if (newStatus == "Approved")
                    {
                        data.SetProperty("U_IsRejected", "No");
                        svc.Update(data);
                        string currentUser = Utilities.Application.Company.UserName;
                        var bpRequest = GetBPDetailsFromUDO(BpCode);
                        CreateBusinessPartner(bpRequest, currentUser);

                    }

                }

            }
            if (newStatus == "Draft")
            {
                Utilities.Application.SBO_Application.MessageBox("Business Partner creation request was *rejected*.\n\n" +"Please review the details, make the necessary corrections, and try again." );
            }
            if (newStatus == "Approved")
            {
                Utilities.Application.SBO_Application.MessageBox($"Business Partner creation request has been *{newStatus}* successfully!\n\n" +"The Business Partner record is now available in the SAP.");
            }
            //Utilities.Application.SBO_Application.MessageBox($"{newStatus} successfully applied.");
            _Form.Close();
            //_Form.Refresh();
        }
        #endregion

        #region GetBPDetailsFromUDO
        private BPRequest GetBPDetailsFromUDO(string bpCode)
        {
            var rs = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            var rs1 = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

            string sql = $@"
                            SELECT ""Code"", ""U_BPName"", ""U_BPType"", ""U_BPPAN"",
                            ""U_PayTerms"", ""U_BPGroup"",
                            ""U_DefPriceList"", ""U_CreditLimit"", ""U_ContactEmail"", ""U_ContactPhone"",
                            ""U_BillToAddr"", ""U_ShipToAddr"", ""U_ContactPerson"", ""U_CommitLimit"", ""U_AcctPayable"", ""U_BankName"", ""U_BankAccNo"",
                            ""U_BankIFSC"", ""U_BankBranch""
                            FROM ""@ABPMASTERH""
                            WHERE ""Code"" = '{bpCode.Replace("'", "''")}'";

            rs.DoQuery(sql);

            string sql1 = $@"SELECT ""U_AttachEntry"" from ""@ABPMASTERC""  WHERE ""Code"" = '{bpCode.Replace("'", "''")}'";
            rs1.DoQuery(sql1);

            if (rs.EoF)
                throw new Exception($"No UDO record found for BPCode {bpCode}");
            return new BPRequest
            {
                BPCode = rs.Fields.Item("Code").Value.ToString(),
                BPName = rs.Fields.Item("U_BPName").Value.ToString(),
                BPPan = rs.Fields.Item("U_BPPAN").Value.ToString(),
                //FederalTaxID = rs.Fields.Item("U_FederalTaxID").Value.ToString(),
                BPType = rs.Fields.Item("U_BPType").Value.ToString(),
                PayTermsGrpCode = Convert.ToInt32(rs.Fields.Item("U_PayTerms").Value),
                //SalesPersonCode = Convert.ToInt32(rs.Fields.Item("U_SalesPersonCode").Value),
                GroupCode = Convert.ToInt32(rs.Fields.Item("U_BPGroup").Value),
                PriceListNum = Convert.ToInt32(rs.Fields.Item("U_DefPriceList").Value),
                CreditLimit = Convert.ToDouble(rs.Fields.Item("U_CreditLimit").Value),
                Email = rs.Fields.Item("U_ContactEmail").Value.ToString(),
                CellPhone = rs.Fields.Item("U_ContactPhone").Value.ToString(),
                BillToAdd = rs.Fields.Item("U_BillToAddr").Value.ToString(),
                ShipToAdd = rs.Fields.Item("U_ShipToAddr").Value.ToString(),
                ContactPerson = rs.Fields.Item("U_ContactPerson").Value.ToString(),
                CommitmentLimit = Convert.ToDouble(rs.Fields.Item("U_CommitLimit").Value),
                AcctPayable = rs.Fields.Item("U_AcctPayable").Value.ToString(),
                BankName = rs.Fields.Item("U_BankName").Value.ToString(),
                BankAccNo = rs.Fields.Item("U_BankAccNo").Value.ToString(),
                BankIFSC = rs.Fields.Item("U_BankIFSC").Value.ToString(),
                BankBranch = rs.Fields.Item("U_BankBranch").Value.ToString()

            };
           
        }
        #endregion

        #region CreateBusinessPartner
        private void CreateBusinessPartner(BPRequest bp, string approvedBy)
        {
            var bpObj = (BusinessPartners)Utilities.Application.Company
            .GetBusinessObject(BoObjectTypes.oBusinessPartners);

            bpObj.CardCode = bp.BPCode;
            bpObj.CardName = bp.BPName;
            bpObj.CardType = bp.BPType == "C" ? BoCardTypes.cCustomer : BoCardTypes.cSupplier;
            bpObj.PayTermsGrpCode = bp.PayTermsGrpCode;
            //bpObj.SalesPersonCode = bp.SalesPersonCode;
            bpObj.GroupCode = bp.GroupCode;
            bpObj.PriceListNum = bp.PriceListNum;
            bpObj.CreditLimit = bp.CreditLimit;
            bpObj.EmailAddress = bp.Email;
            bpObj.Cellular = bp.CellPhone;
            bpObj.DebitorAccount = bp.AcctPayable;
            bpObj.MaxCommitment = bp.CommitmentLimit;
            bpObj.FederalTaxID = bp.BPPan;

               
                // Split full name if possible
                var nameParts = bp.ContactPerson.Split(' ');
                bpObj.ContactEmployees.Name = bp.ContactPerson;
                bpObj.ContactEmployees.FirstName = nameParts[0];
                bpObj.ContactEmployees.LastName = nameParts.Length > 1 ? nameParts[1] : nameParts[0];
                bpObj.ContactEmployees.E_Mail = bp.Email;
                bpObj.ContactEmployees.MobilePhone = bp.CellPhone;
                bpObj.ContactEmployees.Add();
            

            if (!string.IsNullOrWhiteSpace(bp.BillToAdd))
            {
                bpObj.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;
                bpObj.Address = bp.BillToAdd;
                bpObj.Addresses.AddressName = bp.BillToAdd;
                bpObj.Addresses.AddressName2 = bp.BillToAdd;
                bpObj.Addresses.Street = bp.BillToAdd;
                bpObj.Addresses.Add();

            }

            // Shipping Address
            if (!string.IsNullOrWhiteSpace(bp.ShipToAdd))
            {
               
                bpObj.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo;
                bpObj.Address = bp.ShipToAdd;                
                bpObj.Addresses.AddressName = bp.ShipToAdd;
                bpObj.Addresses.AddressName2 = bp.ShipToAdd;
                bpObj.Addresses.Street = bp.ShipToAdd;
                bpObj.Addresses.Add();
            }

            // --- Attachments from UDO Lines ---
            var rsAtt = (Recordset)Utilities.Application.Company
                .GetBusinessObject(BoObjectTypes.BoRecordset);

            string sqlAtt = $@"
            SELECT ""U_AttachFilePath"", ""U_FileName"", ""U_AttachFileType"", ""U_AttachEntry""
            FROM ""@ABPMASTERC""
            WHERE ""Code"" = '{bp.BPCode.Replace("'", "''")}'";

            rsAtt.DoQuery(sqlAtt);

            

            List<int> absEntries = new List<int>();
            while (!rsAtt.EoF)
            {
                absEntries.Add(Convert.ToInt32(rsAtt.Fields.Item("U_AttachEntry").Value));
                rsAtt.MoveNext();
            }

            if (absEntries.Count > 0)
            {
                int mergedAbsEntry = MergeAttachments(absEntries);
                bpObj.AttachmentEntry = mergedAbsEntry;
            }

            //Bank Details
            bpObj.BPBankAccounts.BankCode = bp.BankName;      
            bpObj.BPBankAccounts.Branch = bp.BankBranch;   
            bpObj.BPBankAccounts.AccountNo = bp.BankAccNo;  
            //bpObj.BPBankAccounts.Country = "NP";             
            bpObj.BPBankAccounts.Add();

            bpObj.UserFields.Fields.Item("U_SyncSrc").Value = "Approved BP Creation";

            int res = bpObj.Add();


            if (res != 0)
            {
                string errMsg = Utilities.Application.Company.GetLastErrorDescription();
                Utilities.Application.SBO_Application.MessageBox($"BP creation failed: {errMsg}");
                return;
            }

            var rs = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

            
            string sqlUpdate = $@"
            UPDATE ""@ABPMASTERH""
            SET ""U_SyncStatus""='YES',
                ""U_ApprovedBy""='{approvedBy.Replace("'", "''")}',
                ""U_ApprovedOn""=CURRENT_TIMESTAMP, ""U_ActiveSince""=CURRENT_TIMESTAMP
            WHERE ""Code""='{bp.BPCode.Replace("'", "''")}'";

            rs.DoQuery(sqlUpdate);

            //Utilities.Application.SBO_Application.MessageBox($"BP {bp.BPCode} created and synced successfully!");

        }
        #endregion

        #region BPRequest
        public class BPRequest
        {
            public string BPCode { get; set; }
            public string BPName { get; set; }

            public string BPPan { get; set; }
            public string FederalTaxID { get; set; }
            public string BPType { get; set; }
            public int PayTermsGrpCode { get; set; }
            public int SalesPersonCode { get; set; }
            public int GroupCode { get; set; }
            public int PriceListNum { get; set; }
            public double CreditLimit { get; set; }
            public string Email { get; set; }
            public string CellPhone { get; set; }

            public string BillToAdd { get; set; }
            public string ShipToAdd { get; set; }
            public string ContactPerson { get; set; }
            public double CommitmentLimit { get; set; }
            public string AcctPayable { get; set; }
            public string BankName { get; set; }
            public string BankAccNo { get; set; }
            public string BankIFSC { get; set; }
            public string BankBranch { get; set; }

        }
        #endregion

        #region MergeAttachments
        private int MergeAttachments(List<int> absEntries)
        {
            var newAttachment = (SAPbobsCOM.Attachments2)
                Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2);

            foreach (var absEntry in absEntries)
            {
                var rs = (SAPbobsCOM.Recordset)Utilities.Application.Company
                    .GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                rs.DoQuery($@"SELECT ""srcPath"", ""FileName"", ""FileExt"" 
                      FROM ATC1 WHERE ""AbsEntry"" = {absEntry}");

                while (!rs.EoF)
                {
                    newAttachment.Lines.SourcePath = rs.Fields.Item("srcPath").Value.ToString();
                    newAttachment.Lines.FileName = rs.Fields.Item("FileName").Value.ToString();
                    newAttachment.Lines.FileExtension = rs.Fields.Item("FileExt").Value.ToString();
                    newAttachment.Lines.Add();
                    rs.MoveNext();
                }
            }

            int res = newAttachment.Add();
            if (res != 0)
            {
                string errMsg = Utilities.Application.Company.GetLastErrorDescription();
                throw new Exception($"Attachment merge failed: {errMsg}");
            }

            string newAbsEntry;
            Utilities.Application.Company.GetNewObjectCode(out newAbsEntry);
            return Convert.ToInt32(newAbsEntry);
        }
        #endregion

    }
}

