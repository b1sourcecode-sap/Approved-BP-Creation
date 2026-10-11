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
using System.IO;
using System.Windows.Forms;



namespace Approved_BPCreation
{
    class ClsApprovedBPCreation : Base
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
        string _CardCode = string.Empty,_CardType =string.Empty;
        double _Qty = 0;
        public string ErrMsg = string.Empty;
        public int ErrCode = 0;
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        private static HanaConnection _conn;
        private static HanaDataAdapter _da;
        private static HanaParameter _parm1;
        private static HanaParameter _parm2;
        private static int _errors = 0;
        string strcon, _Sql;
        private SAPbouiCOM.EditText _EditText;
        private SAPbouiCOM.LinkedButton _Linked;
        private SAPbouiCOM.ComboBox _ComboBoxBPType, _ComboBoxPriceList;
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
        private SAPbobsCOM.Recordset _RSPL = null, _RSBPT = null, _RPT = null, _RBP = null, _RSPLG= null, _RPTA = null, _RPBN=null, _RSRej = null;
        private SAPbobsCOM.Recordset  _RSgetAttach;
        private int RowID { get; set; }
        private string _Code { get; set; }
        private string _LineId { get; set; }
        private string _ActivityType = string.Empty;
        string _ItemCode = string.Empty, _ItemName = string.Empty;
        private SAPbobsCOM.Recordset _Rs, _RS, _rs, _RSD, _RSBom, _RSProd, _RSRoute, _RSUsr, _RSMgItm, _RSBomP ,_RSStatus,_RSDoc;
        SAPbouiCOM.ChooseFromListEvent _SysCFLEvent = null;
        SAPbouiCOM.DBDataSource _DBData = null;
        private System.Collections.Generic.List<string> ListOfBlockedID;
        internal string _ServiceID { get; set; }
        internal string _ActivityTypes { get; set; }
        public bool chkStatus = true;
        static int StageID = 0;
        string FileName = string.Empty; string FileType = string.Empty;
        bool hasFinanceAccess = false;
        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsApprovedBPCreation()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //

            // Create and start a permanent STA thread for dialogs
            //FileDialogHelper.Initialize();
            Utilities.Initialize();
        }
        ~ClsApprovedBPCreation()
        {
        }
        #endregion

        #region FormDataEvent
        public override void FormData_Event(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool oBubbleEvent)
        {
            try
            {
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                        _Form.Freeze(true);
                        string status = ((SAPbouiCOM.EditText)_Form.Items.Item("txtStatus").Specific).Value;
                        Disable_Form_Item(status);
                        LoadAccountsPayable();
                        _Form.Freeze(false);
                        break;

                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }

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
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        if (pVal.ItemUID == "mtx_A" && pVal.ColUID == "rowHead" && pVal.Row > 0 && pVal.BeforeAction)
                        {
                            oBubbleEvent = false;
                        }
                        break;

                   
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                        //if (pVal.ItemUID == "1" && (Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                        if (pVal.ItemUID == "btnSubmit" && (Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                        {
                            _CardCode = ((SAPbouiCOM.EditText)_Form.Items.Item("txtBPCode").Specific).Value;
                            _ComboBoxBPType = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbType").Specific;
                            if (((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbType").Specific).Value == "")
                            {
                                Utilities.Message("BP Type must be selected.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                            _CardType = _ComboBoxBPType.Selected.Value;
                            string qryBP = "SELECT COUNT(*) AS Cnt FROM OCRD WHERE \"CardCode\" = '" + _CardCode + "'";
                            Utilities.ExecuteSQL(ref _RBP, qryBP);

                            int count = Convert.ToInt32(_RBP.Fields.Item("Cnt").Value);

                            if (count > 0)
                            {
                                Utilities.Message("Invalid BP Code. Business Partner already exist.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtBPName").Specific).Value == "")
                            {
                                Utilities.Message("BP Name can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                            //if (((SAPbouiCOM.EditText)_Form.Items.Item("txtPanNo").Specific).Value == "")
                            //{
                            //    Utilities.Message("BP Pan No can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            //    oBubbleEvent = false;
                            //    return;
                            //}
                            if (((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbPrice").Specific).Value == "")
                            {
                                Utilities.Message("Default Price must be selected.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                            if (((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbGroup").Specific).Value == "")
                            {
                                Utilities.Message("BP Group must be selected.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtBillTo").Specific).Value == "")
                            {
                                Utilities.Message("Bill To Address can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtShipTo").Specific).Value == "")
                            {
                                Utilities.Message("Ship To Address can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtName").Specific).Value == "")
                            {
                                Utilities.Message("Contact Name can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                            //if (((SAPbouiCOM.EditText)_Form.Items.Item("txtPhone").Specific).Value == "")
                            //{
                            //    Utilities.Message("Contact Number can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            //    oBubbleEvent = false;
                            //    return;
                            //}
                            string phone = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPhone").Specific).Value.Trim();

                            if (string.IsNullOrEmpty(phone))
                            {
                                Utilities.Message("Contact Number cannot be left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                            System.Text.RegularExpressions.Regex phoneRegex = new System.Text.RegularExpressions.Regex(@"^\+?\d{7,15}$");
                            if (!phoneRegex.IsMatch(phone))
                            {
                                Utilities.Message("Please enter a valid contact number (7–15 digits, optional + at start).", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtEmail").Specific).Value == "")
                            {
                                Utilities.Message("Email Address can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }

                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)Form.Items.Item("mtx_A").Specific;
                            if (oMatrix.RowCount == 0 || string.IsNullOrWhiteSpace(((SAPbouiCOM.EditText)oMatrix.Columns.Item("ColFPath").Cells.Item(1).Specific).Value))
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText("Please attach at least one file before adding!",SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                                );

                                oBubbleEvent = false; 
                                return;
                            }
                        }

                        // get status from UDO itself
                        string code = ((SAPbouiCOM.EditText)_Form.Items.Item("txtBPCode").Specific).Value;
                        // or whatever key identifies the record

                        string statusFromDB = string.Empty;

                        SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company
                            .GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        string sql = $@"SELECT ""U_Status"" FROM ""@ABPMASTERH"" WHERE ""Code"" = '{code}'";

                        rs.DoQuery(sql);

                        if (!rs.EoF)
                        {
                            statusFromDB = rs.Fields.Item(0).Value.ToString();
                        }

                        //if (((SAPbouiCOM.EditText)_Form.Items.Item("txtStatus").Specific).Value == "Submitted" && pVal.ItemUID == "1" && (Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) && hasFinanceAccess)
                        if (statusFromDB == "Submitted" && pVal.ItemUID == "1" && (Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) && hasFinanceAccess)
                        {
                            //if (!ValidateFinanceTab())
                            //{
                            //    oBubbleEvent = false; 
                            //    return;
                            //}
                            
                            string credit = ((SAPbouiCOM.EditText)_Form.Items.Item("txtCredit").Specific).Value;
                            if (string.IsNullOrWhiteSpace(credit) || Convert.ToDouble(credit) <= 0)

                            {
                                Utilities.Message("Credit Limit must be greater than zero.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                            string commit = ((SAPbouiCOM.EditText)_Form.Items.Item("txtCmmtL").Specific).Value;
                            if (string.IsNullOrWhiteSpace(commit) || Convert.ToDouble(commit) <= 0)
                            {
                                Utilities.Message("Commitment Limit must be greater than zero.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }

                            double credit1 = Convert.ToDouble(credit);
                            double commit1 = Convert.ToDouble(commit);

                           
                            if (credit1 > commit1)
                            {
                                Utilities.Message("Credit Limit cannot be greater than Commitment Limit.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                            //if (((SAPbouiCOM.ComboBox)_Form.Items.Item("txtBankN").Specific).Value == "")
                            //{
                            //    Utilities.Message("Bank Name can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            //    oBubbleEvent = false;
                            //    return;
                            //}
                            //if (((SAPbouiCOM.EditText)_Form.Items.Item("txtAcNo").Specific).Value == "")
                            //{
                            //    Utilities.Message("Account Number can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            //    oBubbleEvent = false;
                            //    return;
                            //}
                            //if (((SAPbouiCOM.EditText)_Form.Items.Item("txtIFSC").Specific).Value == "")
                            //{
                            //    Utilities.Message("Bank IFSC can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            //    oBubbleEvent = false;
                            //    return;
                            //}
                            //if (((SAPbouiCOM.EditText)_Form.Items.Item("txtBBranch").Specific).Value == "")
                            //{
                            //    Utilities.Message("Bank Branch can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            //    oBubbleEvent = false;
                            //    return;
                            //}
                            if (((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbPayT").Specific).Value == "")
                            {
                                Utilities.Message("Payment Terms must be selected.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                            if (((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbAcct").Specific).Value == "")
                            {
                                Utilities.Message("Account Payable must be selected.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }


                            //if (Validation() == false)
                            //{
                            //    oBubbleEvent = false;
                            //    return;
                            //}
                            errorFlag = false;
                            //AddPO();
                            if (errorFlag == true)
                            {
                                oBubbleEvent = false;
                                return;
                            }
                        }
                        if (pVal.ItemUID == "1" && (Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                        {
                            string departmentName = GetCurrentUserDepartment();

                            //if (departmentName.Equals("Account", StringComparison.OrdinalIgnoreCase))
                            //{
                            //    Utilities.Message("Users from the Finance Department are not allowed to create new BP Creation requests.",SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                            //    oBubbleEvent = false;
                            //    return;
                            //}

                            var restrictedDepartments = new[]
                                {
                                    "Account",
                                    "Accounts",
                                    "Finance",
                                    "Finance Department",
                                    "Accounting",
                                    "Accounts & Finance"
                              };

                            if (restrictedDepartments.Any(d =>
                                    string.Equals(d, departmentName?.Trim(), StringComparison.OrdinalIgnoreCase)))
                            {
                                Utilities.Message(
                                    "Users from the Finance Department are not allowed to create new BP Creation requests.",
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                                oBubbleEvent = false;
                                return;
                            }

                        }

                        
                        break;

                }
            }

            #endregion

            #region After Action
            else
            {

                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (!pVal.BeforeAction) {

                            if (pVal.ItemUID == "mtx_A")

                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtx_A").Specific;
                                // Select clicked row
                                oMatrix.SelectRow(pVal.Row, true, false);


                                // Optionally set focus
                                //oMatrix.Columns.Item(1).Cells.Item(pVal.Row).Click();
                            }
                      
                        }
                       break;


                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:

                        //if (pVal.ColUID == "V_8" && (Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
                        //{
                        //    _ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific;
                        //    if (_ComboBoxC.Selected.Value == "Route Stage")
                        //    {
                        //        _Matrix.CommonSetting.SetCellFontSize(pVal.Row, 1, 12);
                        //        _Matrix.CommonSetting.SetCellFontSize(pVal.Row, 2, 12);
                        //        _Matrix.CommonSetting.SetCellFontSize(pVal.Row, 3, 12);
                        //        _Matrix.CommonSetting.SetCellFontStyle(pVal.Row, 1, BoFontStyle.fs_Bold);
                        //        _Matrix.CommonSetting.SetCellFontStyle(pVal.Row, 2, BoFontStyle.fs_Bold);
                        //        _Matrix.CommonSetting.SetCellFontStyle(pVal.Row, 3, BoFontStyle.fs_Bold);
                        //        //SetCellForeColor(pVal.Row, 1, greenColor);
                        //        _Matrix.CommonSetting.SetCellEditable(pVal.Row, 4, false);
                        //        _Matrix.CommonSetting.SetCellEditable(pVal.Row, 6, false);
                        //        _Matrix.CommonSetting.SetCellEditable(pVal.Row, 7, false);
                        //    }

                        //}
                        if (pVal.ItemUID == "cmbPList")
                        {
                            //_ComboBoxM = (SAPbouiCOM.ComboBox)Form.Items.Item("cmbPList").Specific;
                            //_EditText = (SAPbouiCOM.EditText)Form.Items.Item("txtPNo").Specific;
                            //_EditText.Value = _ComboBoxM.Selected.Description.ToString();
                        }
                        if (pVal.ItemUID == "cmbType" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT && pVal.BeforeAction == false)
                        {
                            //BP GroupCode
                            SAPbouiCOM.ComboBox cmbType = (SAPbouiCOM.ComboBox)Form.Items.Item("cmbType").Specific;
                            string selectedType = cmbType.Selected.Value; 

                            SAPbouiCOM.ComboBox cmbGroup = (SAPbouiCOM.ComboBox)Form.Items.Item("cmbGroup").Specific;

                            while (cmbGroup.ValidValues.Count > 0)
                            {
                                cmbGroup.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            }

                            string qryGroup = $@"SELECT ""GroupCode"", ""GroupName"" FROM OCRG WHERE ""GroupType"" = '{selectedType}' ORDER BY ""GroupName"" ASC";

                            Utilities.ExecuteSQL(ref _RSPLG, qryGroup);

                            while (!_RSPLG.EoF)
                            {
                                string groupCode = _RSPLG.Fields.Item("GroupCode").Value.ToString();
                                string groupName = _RSPLG.Fields.Item("GroupName").Value.ToString();

                                cmbGroup.ValidValues.Add(groupCode, groupName);
                                cmbGroup.Item.DisplayDesc = true;

                                _RSPLG.MoveNext();
                            }

                        }

                        break;

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                        //if(pVal.ItemUID == "btnSubmit" && Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        //{
                        //    //_Form.Items.Item("txtStatus").Enabled = true;
                        //   ((SAPbouiCOM.EditText)Form.Items.Item("txtStatus").Specific).Value = "Submitted";
                        //    if (Form.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        //         Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        //    _Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                        //    string Code = ((SAPbouiCOM.EditText)_Form.Items.Item("txtBPCode").Specific).Value;
                        //    Utilities.ExecuteSQL(ref _RSUsr, "Select T0.\"USER_CODE\" " +
                        //    "From \"OUSR\" T0 where T0.\"U_FinanceA\" ='Yes'");
                        //    List<string> approvers = new List<string>();
                        //    while (!_RSUsr.EoF)
                        //    {
                        //        string code = _RSUsr.Fields.Item(0).Value.ToString().Trim();
                        //        if (!string.IsNullOrEmpty(code)) approvers.Add(code);
                        //        _RSUsr.MoveNext();
                        //    }

                        //    if (approvers.Count > 0)
                        //        SendInternalAlert(Code, approvers.Count == 1 ? new List<string> { approvers[0] } : approvers,"Submit");
                        //    else
                        //        Utilities.Application.SBO_Application.MessageBox("No approvers found.");
                        //    Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                        //}
                        if (pVal.ItemUID == "btnSubmit")
                        {
                            if (Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                ((SAPbouiCOM.EditText)Form.Items.Item("txtStatus").Specific).Value = "Submitted";

                                if (Form.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                    Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

                                Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                                string Code = ((SAPbouiCOM.EditText)Form.Items.Item("txtBPCode").Specific).Value;

                                Utilities.ExecuteSQL(ref _RSUsr, "SELECT T0.\"USER_CODE\" FROM \"OUSR\" T0 WHERE T0.\"U_FinanceA\" = 'Yes'");

                                List<string> approvers = new List<string>();
                                while (!_RSUsr.EoF)
                                {
                                    string code = _RSUsr.Fields.Item(0).Value.ToString().Trim();
                                    if (!string.IsNullOrEmpty(code)) approvers.Add(code);
                                    _RSUsr.MoveNext();
                                }

                                //if (approvers.Count > 0)
                                //    SendInternalAlert(Code, approvers.Count == 1 ? new List<string> { approvers[0] } : approvers, "Submit");
                                //else
                                //    Utilities.Application.SBO_Application.MessageBox("No approvers found.");

                                Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                            else
                            {
                                Utilities.Application.SBO_Application.MessageBox("The form is not ready for submission. Please update or save before submitting.");
                                return;
                            }
                        }



                        if (pVal.ItemUID == "btnAproved" && Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            ((SAPbouiCOM.EditText)_Form.Items.Item("txtStatus").Specific).Value = "Awaiting Approval";
                            _Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            _Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                            //after finance update alerts to approver
                            string Code1 = ((SAPbouiCOM.EditText)_Form.Items.Item("txtBPCode").Specific).Value;
                            Utilities.ExecuteSQL(ref _RSUsr, "Select T0.\"USER_CODE\" " +
                          "From \"OUSR\" T0 where T0.\"U_BPAuth\" ='Yes'");
                            List<string> approvers = new List<string>();
                            while (!_RSUsr.EoF)
                            {
                                string code = _RSUsr.Fields.Item(0).Value.ToString().Trim();
                                if (!string.IsNullOrEmpty(code)) approvers.Add(code);
                                _RSUsr.MoveNext();
                            }

                            //if (approvers.Count > 0)
                            //    SendInternalAlert(Code1, approvers.Count == 1 ? new List<string> { approvers[0] } : approvers,"Approve");
                            //else
                            //    Utilities.Application.SBO_Application.MessageBox("No approvers found.");
                            Form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);


                        }

                        #region Add Attachment Path 
                        if (pVal.ItemUID == "btnBrowse")
                        {
                            // string filePath = ShowFileDialog();
                            //string filePath = FileDialogHelper.ShowFileDialog("All files (*.*)|*.*", "Select a File");
                            string filePath = Utilities.ShowFileDialog("All files (*.*)|*.*", "Select a File");

                            if (!string.IsNullOrEmpty(filePath))
                            {
                                AttachmentResult result = SaveAttachmentToSAP(filePath);
                                if (result.AbsEntry != -1)
                                {
                                    string destPath = result.DestPath;
                                    int absEntry = result.AbsEntry;
                                    SAPbouiCOM.Form oForm = Utilities.Application.SBO_Application.Forms.ActiveForm;
                                    _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtx_A").Specific;
                                    _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@ABPMASTERC");

                                    if (_ChildDBDataSource.Size > 0)
                                    {
                                        int lastIndex = _ChildDBDataSource.Size - 1;
                                        string lastPath = _ChildDBDataSource.GetValue("U_AttachFilePath", lastIndex)?.Trim();
                                        string lastFile = _ChildDBDataSource.GetValue("U_FileName", lastIndex)?.Trim();
                                        if (string.IsNullOrEmpty(lastPath) && string.IsNullOrEmpty(lastFile))
                                        {
                                            _ChildDBDataSource.RemoveRecord(lastIndex);
                                        }
                                    }

                                    int newRow = _ChildDBDataSource.Size;
                                    _ChildDBDataSource.InsertRecord(newRow);

                                    _ChildDBDataSource.SetValue("U_AttachFilePath", newRow, destPath);
                                    _ChildDBDataSource.SetValue("U_AttachEntry", newRow, absEntry.ToString());
                                    _ChildDBDataSource.SetValue("U_AttachDate", newRow, DateTime.Now.ToString("yyyyMMdd"));
                                    _ChildDBDataSource.SetValue("U_FileName", newRow, FileName);
                                    _ChildDBDataSource.SetValue("U_AttachFileType", newRow, FileType);
                                    _ChildDBDataSource.SetValue("LineId", newRow, (newRow + 1).ToString());

                                    _Matrix.LoadFromDataSource();

                                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                    {
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

                                        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    }
                                       

                                    Utilities.Application.SBO_Application.StatusBar.SetText("Attachment added and UDO updated.",BoMessageTime.bmt_Short,BoStatusBarMessageType.smt_Success);
                                }
                            }
                        }
                        #endregion

                        #region View Attachment Path
                        //if (pVal.ItemUID == "btnDisplay" && Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        if (pVal.ItemUID == "btnDisplay")
                            {
                            
                            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtx_A").Specific;

                            int selectedRow = _Matrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);

                            if (selectedRow > 0) // valid row selected
                            {
                                string absEntryStr = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("Col_Entry").Cells.Item(selectedRow).Specific).Value;


                                if (int.TryParse(absEntryStr, out int absEntry))
                                {
                                    OpenAttachment(absEntry);
                                }
                                else
                                {
                                    Utilities.Application.SBO_Application.StatusBar.SetText("Invalid AbsEntry in selected row.",SAPbouiCOM.BoMessageTime.bmt_Short,SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                }
                            }
                            else
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText(
                                    "Please select a row in the matrix.",SAPbouiCOM.BoMessageTime.bmt_Short,SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            }
                        }
                        #endregion

                        #region Delete Attachment
                        if (pVal.ItemUID == "btnDelete")
                        {
                            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtx_A").Specific;
                            _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@ABPMASTERC");

                            int selectedRow = _Matrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);

                            if (selectedRow > 0) 
                            {
                                int rowIndex = selectedRow - 1; 

                                try
                                {
                                   
                                    _ChildDBDataSource.RemoveRecord(rowIndex);

                                  
                                    for (int i = 0; i < _ChildDBDataSource.Size; i++)
                                    {
                                        _ChildDBDataSource.SetValue("LineId", i, (i + 1).ToString());

                                    }

                                    // Sync with matrix
                                    _Matrix.LoadFromDataSource();

                                    Utilities.Application.SBO_Application.StatusBar.SetText($"Attachment at row {selectedRow} deleted and rows resequenced.",SAPbouiCOM.BoMessageTime.bmt_Short,SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                }
                                catch (Exception ex)
                                {
                                    Utilities.Application.SBO_Application.StatusBar.SetText($"Error while deleting row: {ex.Message}",SAPbouiCOM.BoMessageTime.bmt_Short,SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                }
                            }
                            else
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText("Please select a row in the matrix to delete.",SAPbouiCOM.BoMessageTime.bmt_Short,SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            }
                        }
                        #endregion
                        break;

                    #region Set CFL Values
                    //case BoEventTypes.et_CHOOSE_FROM_LIST:
                    //    _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
                    //    if (pVal.ItemUID == "txtPrNo" && _SysCFLEvent.SelectedObjects != null)
                    //    {
                    //        // Set main DB fields
                    //        string itemCode = _SysCFLEvent.SelectedObjects.GetValue("ItemCode", 0).ToString().Trim();
                    //        _MasterDBDataSource.SetValue("U_PRNum", 0, itemCode);
                    //        _MasterDBDataSource.SetValue("U_PRName", 0, _SysCFLEvent.SelectedObjects.GetValue("ItemName", 0).ToString().Trim());
                    //        _MasterDBDataSource.SetValue("U_UOM", 0, _SysCFLEvent.SelectedObjects.GetValue("InvntryUom", 0).ToString().Trim());
                    //        //_MasterDBDataSource.SetValue("U_PPNum", 0, Utilities.getMaxColumnValueNumWithItem("@PRODUCTIONPH", "U_PPNum", itemCode));
                    //        _MasterDBDataSource.SetValue("U_PQty", 0, "1");
                    //        string whsCode = "";
                    //        SAPbobsCOM.Recordset oRSWhse = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //        string query = $"SELECT \"ToWH\" FROM \"OITT\" WHERE \"Code\" = '{itemCode}'";

                    //        oRSWhse.DoQuery(query);

                    //        if (!oRSWhse.EoF)
                    //        {
                    //            _MasterDBDataSource.SetValue("U_Whse", 0, oRSWhse.Fields.Item("ToWH").Value.ToString());
                    //        }
                    //        string orderType = ((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbType").Specific).Value;
                    //        _Form.Select();
                    //        LoadProductionPC(orderType, itemCode);
                    //        AddSequenceNumbersToMatrix(_Matrix);
                    //        //if(orderType =="Special")
                    //        //{
                    //        //    _Matrix.AddRow(1, -1);
                    //        //}
                    //        AddSequenceNumbersToMatrix(_Matrix);
                    //    }
                    //    if (pVal.ItemUID == "txtWhse" && _SysCFLEvent.SelectedObjects != null)
                    //    {
                    //        _MasterDBDataSource.SetValue("U_Whse", 0, _SysCFLEvent.SelectedObjects.GetValue("WhsCode", 0).ToString().Trim());
                    //    }
                    //    if (pVal.ItemUID == "27" && pVal.ColUID == "V_0" && _SysCFLEvent.SelectedObjects != null)
                    //    {

                    //        try
                    //        {
                    //            _ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific;
                    //            if (_ComboBoxC.Selected.Value == "Item")
                    //            {
                    //                _Matrix.FlushToDataSource();
                    //                _ChildDBDataSource.SetValue("U_ItemCode", _ChildDBDataSource.Offset, _SysCFLEvent.SelectedObjects.GetValue("ItemCode", 0).ToString().Trim());
                    //                _ChildDBDataSource.SetValue("U_ItemName", _ChildDBDataSource.Offset, _SysCFLEvent.SelectedObjects.GetValue("ItemName", 0).ToString().Trim());
                    //                _ChildDBDataSource.SetValue("U_UOM", _ChildDBDataSource.Offset, _SysCFLEvent.SelectedObjects.GetValue("InvntryUom", 0).ToString().Trim());
                    //                _ChildDBDataSource.SetValue("U_Quantity", _ChildDBDataSource.Offset, "1");
                    //                _ComboBoxB = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_9").Cells.Item(pVal.Row).Specific;
                    //                _Matrix.LoadFromDataSource();
                    //                if (_SysCFLEvent.SelectedObjects.GetValue("IssueMthd", 0).ToString().Trim() == "B")
                    //                {
                    //                    _ComboBoxB.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    //                }
                    //                else
                    //                {
                    //                    _ComboBoxB.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
                    //                }
                    //                _Matrix.FlushToDataSource();
                    //                _Matrix.LoadFromDataSource();
                    //                AddNewRow();
                    //                _ComboBoxCRow = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row + 1).Specific;
                    //                _ComboBoxCRow.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    //            }
                    //        }
                    //        catch
                    //        {
                    //            Utilities.Application.Company.GetLastError(out ErrCode, out ErrMsg);
                    //        }

                    //    }
                    //    if (pVal.ItemUID == "27" && pVal.ColUID == "V_3" && _SysCFLEvent.SelectedObjects != null)
                    //    {
                    //        try
                    //        {
                    //            _Matrix.FlushToDataSource();
                    //            _ChildDBDataSource.Offset = pVal.Row - 1; // Set the correct row
                    //            _ChildDBDataSource.SetValue("U_Whse", _ChildDBDataSource.Offset, _SysCFLEvent.SelectedObjects.GetValue("WhsCode", 0).ToString().Trim());
                    //            _Matrix.LoadFromDataSource();
                    //            _Matrix.FlushToDataSource();
                    //            //int i = pVal.Row;
                    //            //dynamic cflArg1 = pVal;
                    //            //string UID1 = cflArg1.ChooseFromListUID;
                    //            //dynamic dt1 = cflArg1.SelectedObjects;
                    //            //_Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("27").Specific;
                    //            //((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(i).Specific).Value = dt1.GetValue("WhsCode", 0).ToString();
                    //        }

                    //        catch (Exception ex)
                    //        {
                    //            // Handle the exception, log it, or display an error message.
                    //        }
                    //        chkStatus = true;
                    //    }
                    //    break;
                    //#endregion

                    //#region Key Down 
                    //case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                    //    if(pVal.ItemUID == "txtQty"  && !pVal.BeforeAction)
                    //        {
                    //        double inputQty = 0;
                    //        double.TryParse(((SAPbouiCOM.EditText)_Form.Items.Item("txtQty").Specific).Value, out inputQty);

                    //        var childDS = _Form.DataSources.DBDataSources.Item("@PRODUCTIONPC");
                    //        for (int i = 0, len = childDS.Size; i < len; i++)
                    //        {
                    //            double baseQty = 0;
                    //            double.TryParse(childDS.GetValue("U_BaseQty", i).ToString(), out baseQty);

                    //            childDS.SetValue("U_PlannedQty", i, (baseQty * inputQty).ToString());
                    //        }
                    //        _Matrix.LoadFromDataSource();
                    //        // ((SAPbouiCOM.Grid)_Form.Items.Item("yourGridUID").Specific).LoadFromDataSource();
                    //    }
                    //    break;
                    #endregion
                    
                }
            }

            #endregion
        }
        #endregion

        #region MENU EVENT
        public override void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool oBubbleEvent)
        {
            #region BEFORE ACTION
            if (pVal.BeforeAction)
            {
                Form.Freeze(true);
                switch (pVal.MenuUID)
                {
                    case Constants.System_Menus.mnu_Duplicate:
                        {
                            _PRNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value;
                            //Disable_Form_Item();
                        }
                        break;
                }
                Form.Freeze(false);
            }
            #endregion

            #region AFTER ACTION
            else
            {
                try
                {
                    Form.Freeze(true);
                    switch (pVal.MenuUID)
                    {
                        case Constants.System_Menus.mnu_FIND:
                            //_Form.Items.Item("txtPPNo").Enabled = true;
                            _Form.Items.Item("txtBPCode").Enabled = true;
                            break;
                        case Constants.System_Menus.mnu_ADD:
                                EnableForAddMode();
                      
                            break;
                        case Constants.System_Menus.mnu_Duplicate:
                            {
                                //_MasterDBDataSource.SetValue("U_PPNum", 0, (Utilities.getMaxColumnValueNumWithItem("@PRODUCTIONPH", "U_PPNum", _PRNum)));
                                //_Form.Items.Item("txtPPNo").Enabled = false;
                                //_ComboBoxH.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                //AddSequenceNumbersToMatrix(_Matrix);
                            }
                            break;
                        case Constants.System_Menus.mnu_FIRST:
                            if (_Form.Mode != BoFormMode.fm_OK_MODE)
                            {
                                // Cancel the changes (like hitting "No")
                                _Form.Mode = BoFormMode.fm_OK_MODE; // Reset mode to OK
                            }
                            break;

                        case Constants.System_Menus.mnu_LAST:
                            //_Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("27").Specific;
                            //AddSequenceNumbersToMatrix(_Matrix);
                            //Disable_Form_item();

                            if (_Form.Mode != BoFormMode.fm_OK_MODE)
                            {
                                // Cancel the changes (like hitting "No")
                                _Form.Mode = BoFormMode.fm_OK_MODE; // Reset mode to OK
                            }
                            break;

                        case Constants.System_Menus.mnu_NEXT:
                            if (_Form.Mode != BoFormMode.fm_OK_MODE)
                            {
                                // Cancel the changes (like hitting "No")
                                _Form.Mode = BoFormMode.fm_OK_MODE; // Reset mode to OK
                            }

                            break;

                        case Constants.System_Menus.mnu_PREVIOUS:
                            if (_Form.Mode != BoFormMode.fm_OK_MODE)
                            {
                                // Cancel the changes (like hitting "No")
                                _Form.Mode = BoFormMode.fm_OK_MODE; // Reset mode to OK
                            }
                            break;

                    }
                    if (pVal.MenuUID == "1292" && pVal.BeforeAction == false)
                    {
                        AddNewRow();
                    }
                }

                catch (Exception ex)
                {
                    Utilities.ShowErrorMessage(ex.Message);
                }
                finally { Form.Freeze(false); }
            }
            #endregion
        }
        #endregion

        #region CHOOSE FROM LIST EVENT

        public override void ChooseFromList_Event()
        {
            try
            {
                _Form.Freeze(true);
                if (this._CFLEvent.ItemUID == "27" && _CFLEvent.ColUID == "V_0")
                {
                    for (int i = 0; i < _CFLEvent.oCFL_SelectedResult.Rows.Count + _CFLEvent.Row - 1; i++)
                    {
                        _ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(_Matrix.VisualRowCount).Specific;
                        _ComboBoxB = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_9").Cells.Item(_Matrix.VisualRowCount).Specific;
                        if (_ComboBoxC.Selected.Value == "Resource")
                        {
                            int a = i;
                            _Matrix.FlushToDataSource();
                            _ChildDBDataSource.SetValue("U_ItemCode", _ChildDBDataSource.Offset, _CFLEvent.oCFL_SelectedResult.Rows[i][0].ToString());
                            _ChildDBDataSource.SetValue("U_ItemName", _ChildDBDataSource.Offset, _CFLEvent.oCFL_SelectedResult.Rows[i][1].ToString());
                            _ChildDBDataSource.SetValue("U_UOM", _ChildDBDataSource.Offset, _CFLEvent.oCFL_SelectedResult.Rows[i][2].ToString());
                            _ChildDBDataSource.SetValue("U_Quantity", _ChildDBDataSource.Offset, "1");
                            _ComboBoxB.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            _Matrix.LoadFromDataSource();
                            AddNewRow();
                            a = i;
                            _ComboBoxCRow = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(_Matrix.VisualRowCount).Specific;
                            _ComboBoxCRow.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        else
                        {
                            int a = i;
                            _Matrix.FlushToDataSource();
                            _ChildDBDataSource.SetValue("U_ItemCode", _ChildDBDataSource.Offset, _CFLEvent.oCFL_SelectedResult.Rows[i][0].ToString());
                            _ChildDBDataSource.SetValue("U_ItemName", _ChildDBDataSource.Offset, _CFLEvent.oCFL_SelectedResult.Rows[i][1].ToString());
                            _Matrix.LoadFromDataSource();
                            AddNewRow();
                            a = i;
                            _ComboBoxCRow = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(_Matrix.VisualRowCount).Specific;
                            _ComboBoxCRow.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
            finally
            {
                _Form.Freeze(false);
            }
        }

        #endregion

        #region Form Default
        //internal void FormDefault()
        //{
        //    try
        //    {
        //        string departmentName = GetCurrentUserDepartment();

        //        _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@ABPMASTERH");
        //        _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@ABPMASTERC");
        //        _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtx_A").Specific;

        //        if (departmentName.Equals("FINANCE & ACCOUNTS", StringComparison.OrdinalIgnoreCase))
        //        {
        //            _Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
        //            return;
        //        }
        //        _Form.Freeze(true);
        //        //_Form.DataBrowser.BrowseBy = "txtEntry";
        //        _Form.Settings.Enabled = true;
        //        UserObjectsMD oUserObjectMD = (UserObjectsMD)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
        //        _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("mtx_A").Specific;
        //        _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@ABPMASTERH");
        //        _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@ABPMASTERC");
        //        ((SAPbouiCOM.EditText)_Form.Items.Item("txtCrtBy").Specific).Value = Utilities.Application.SBO_Application.Company.UserName;
        //        ((SAPbouiCOM.EditText)_Form.Items.Item("txtCrtOn").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
        //        ((SAPbouiCOM.EditText)_Form.Items.Item("txtStatus").Specific).Value = "Draft";

        //        //Price List
        //        SAPbouiCOM.ComboBox cmbPrice = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbPrice").Specific;

        //        // Clear old values
        //        ClearValidValues(cmbPrice.ValidValues);

        //        // Reload from query
        //        string priceList = "SELECT \"ListNum\",\"ListName\" FROM OPLN ORDER BY \"ListName\" ASC";
        //        Utilities.ExecuteSQL(ref _RSPL, priceList);

        //        while (!_RSPL.EoF)
        //        {
        //            string listId = _RSPL.Fields.Item("ListNum").Value.ToString();
        //            string listName = _RSPL.Fields.Item("ListName").Value.ToString();

        //            cmbPrice.ValidValues.Add(listId, listName);

        //            _RSPL.MoveNext();
        //        }

        //        cmbPrice.Item.DisplayDesc = true;



        //        //Header CardType
        //        SAPbouiCOM.ComboBox cmbType = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbType").Specific;
        //        ClearValidValues(cmbType.ValidValues);

        //        string qryCardType = "SELECT DISTINCT \"CardType\", " +
        //                             "CASE WHEN \"CardType\" = 'C' THEN 'Customer' " +
        //                             "     WHEN \"CardType\" = 'S' THEN 'Vendor' " +
        //                             "END AS \"CardDesc\" FROM OCRD";

        //        Utilities.ExecuteSQL(ref _RSBPT, qryCardType);

        //        while (!_RSBPT.EoF)
        //        {
        //            string cardType = _RSBPT.Fields.Item("CardType").Value.ToString();
        //            string cardDesc = _RSBPT.Fields.Item("CardDesc").Value.ToString();

        //            cmbType.ValidValues.Add(cardType, cardDesc);

        //            _RSBPT.MoveNext();
        //        }
        //        cmbType.Item.DisplayDesc = true;


        //        //Header BPGroup
        //        SAPbouiCOM.ComboBox cmbGroup = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbGroup").Specific;
        //        ClearValidValues(cmbGroup.ValidValues);

        //        string group = "SELECT DISTINCT \"GroupCode\", \"GroupName\" FROM OCRG ORDER BY \"GroupName\" ASC";
        //        Utilities.ExecuteSQL(ref _RSPLG, group);

        //        while (!_RSPLG.EoF)
        //        {
        //            string groupCode = _RSPLG.Fields.Item("GroupCode").Value.ToString();
        //            string groupName = _RSPLG.Fields.Item("GroupName").Value.ToString();

        //            cmbGroup.ValidValues.Add(groupCode, groupName);

        //            _RSPLG.MoveNext();
        //        }
        //        cmbGroup.Item.DisplayDesc = true;



        //        // ===== Finance Payment Terms =====
        //        SAPbouiCOM.ComboBox cmbPayT = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbPayT").Specific;
        //        ClearValidValues(cmbPayT.ValidValues);
        //        string qryPaymentGroup = "SELECT \"GroupNum\", \"PymntGroup\" FROM OCTG ORDER BY \"PymntGroup\" ASC";
        //        Utilities.ExecuteSQL(ref _RPT, qryPaymentGroup);
        //        while (!_RPT.EoF)
        //        {
        //            cmbPayT.ValidValues.Add(_RPT.Fields.Item("GroupNum").Value.ToString(),
        //                                    _RPT.Fields.Item("PymntGroup").Value.ToString());
        //            _RPT.MoveNext();
        //        }
        //        cmbPayT.Item.DisplayDesc = true;

        //        // ===== Finance Account Payable =====
        //        SAPbouiCOM.ComboBox cmbAcct = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbAcct").Specific;
        //        ClearValidValues(cmbAcct.ValidValues);
        //        string accPayable = "SELECT DISTINCT \"AcctCode\", \"AcctName\" FROM OACT WHERE \"LocManTran\" = 'Y'";
        //        Utilities.ExecuteSQL(ref _RPTA, accPayable);
        //        while (!_RPTA.EoF)
        //        {
        //            cmbAcct.ValidValues.Add(_RPTA.Fields.Item("AcctCode").Value.ToString(),
        //                                    _RPTA.Fields.Item("AcctName").Value.ToString());
        //            _RPTA.MoveNext();
        //        }
        //        cmbAcct.Item.DisplayDesc = true;

        //        // ===== Finance Bank Name =====
        //        SAPbouiCOM.ComboBox cmbBank = (SAPbouiCOM.ComboBox)_Form.Items.Item("txtBankN").Specific;
        //        ClearValidValues(cmbBank.ValidValues);
        //        string bankName = "SELECT DISTINCT \"BankCode\", \"BankName\" FROM ODSC";
        //        Utilities.ExecuteSQL(ref _RPBN, bankName);
        //        while (!_RPBN.EoF)
        //        {
        //            try
        //            {
        //                cmbBank.ValidValues.Add(_RPBN.Fields.Item("BankCode").Value.ToString(),
        //                                        _RPBN.Fields.Item("BankName").Value.ToString());
        //            }
        //            catch { /* Skip duplicates */ }
        //            _RPBN.MoveNext();
        //        }
        //        cmbBank.Item.DisplayDesc = true;



        //    }
        //    catch (Exception ex)
        //    {
        //        Utilities.ShowErrorMessage(ex.Message);
        //    }
        //    finally
        //    {
        //        _Form.Freeze(false);
        //        _Form.Visible = true;
        //    }
        //    _Form.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //    //_Form.Items.Item("btnSubmit").Enabled = false;
        //    _Form.Items.Item("txtStatus").Enabled = false;
        //    _Form.Items.Item("txtCrtBy").Enabled = false;
        //    _Form.Items.Item("txtCrtOn").Enabled = false;
        //    _Form.Items.Item("txtAppBy").Enabled = false;
        //    _Form.Items.Item("txtAppOn").Enabled = false;

        //    _Form.Freeze(false);
        //    _Form.PaneLevel = 1;
        //}
        #endregion 

        internal void FormDefault()
        {
            try
            {
                _Form.Freeze(true);

                string departmentName = GetCurrentUserDepartment();
                //bool isFinance = departmentName.Equals("ACCOUNTS", StringComparison.OrdinalIgnoreCase);

                bool isFinance =
                departmentName.Equals("ACCOUNTS", StringComparison.OrdinalIgnoreCase) ||
                departmentName.Equals("ACCOUNT", StringComparison.OrdinalIgnoreCase) ||
                departmentName.Equals("FINANCE", StringComparison.OrdinalIgnoreCase) ||
                departmentName.Equals("FINANCE DEPARTMENT", StringComparison.OrdinalIgnoreCase) ||
                departmentName.Equals("ACCOUNTING", StringComparison.OrdinalIgnoreCase) ||
                departmentName.Equals("ACCOUNTS & FINANCE", StringComparison.OrdinalIgnoreCase);



                _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@ABPMASTERH");
                _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@ABPMASTERC");
                _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtx_A").Specific;

                _Form.Settings.Enabled = true;
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtCrtBy").Specific).Value = Utilities.Application.SBO_Application.Company.UserName;
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtCrtOn").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtStatus").Specific).Value = "Draft";

                // === Price List Combo ===
                SAPbouiCOM.ComboBox cmbPrice = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbPrice").Specific;
                ClearValidValues(cmbPrice.ValidValues);

                string priceList = "SELECT \"ListNum\",\"ListName\" FROM OPLN ORDER BY \"ListName\" ASC";
                Utilities.ExecuteSQL(ref _RSPL, priceList);
                while (!_RSPL.EoF)
                {
                    cmbPrice.ValidValues.Add(_RSPL.Fields.Item("ListNum").Value.ToString(),
                                             _RSPL.Fields.Item("ListName").Value.ToString());
                    _RSPL.MoveNext();
                }
                cmbPrice.Item.DisplayDesc = true;

                // === Card Type Combo ===
                SAPbouiCOM.ComboBox cmbType = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbType").Specific;
                ClearValidValues(cmbType.ValidValues);

                string qryCardType = @"SELECT DISTINCT ""CardType"",
                                CASE WHEN ""CardType"" = 'C' THEN 'Customer'
                                     WHEN ""CardType"" = 'S' THEN 'Vendor' END AS ""CardDesc""
                                FROM OCRD";
                Utilities.ExecuteSQL(ref _RSBPT, qryCardType);
                while (!_RSBPT.EoF)
                {
                    cmbType.ValidValues.Add(_RSBPT.Fields.Item("CardType").Value.ToString(),
                                            _RSBPT.Fields.Item("CardDesc").Value.ToString());
                    _RSBPT.MoveNext();
                }
                cmbType.Item.DisplayDesc = true;

                // === BP Group Combo ===
                SAPbouiCOM.ComboBox cmbGroup = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbGroup").Specific;
                ClearValidValues(cmbGroup.ValidValues);

                string group = "SELECT DISTINCT \"GroupCode\", \"GroupName\" FROM OCRG ORDER BY \"GroupName\" ASC";
                Utilities.ExecuteSQL(ref _RSPLG, group);
                while (!_RSPLG.EoF)
                {
                    cmbGroup.ValidValues.Add(_RSPLG.Fields.Item("GroupCode").Value.ToString(),
                                             _RSPLG.Fields.Item("GroupName").Value.ToString());
                    _RSPLG.MoveNext();
                }
                cmbGroup.Item.DisplayDesc = true;

                // === Payment Terms ===
                SAPbouiCOM.ComboBox cmbPayT = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbPayT").Specific;
                ClearValidValues(cmbPayT.ValidValues);

                string qryPaymentGroup = "SELECT \"GroupNum\", \"PymntGroup\" FROM OCTG ORDER BY \"PymntGroup\" ASC";
                Utilities.ExecuteSQL(ref _RPT, qryPaymentGroup);
                while (!_RPT.EoF)
                {
                    cmbPayT.ValidValues.Add(_RPT.Fields.Item("GroupNum").Value.ToString(),
                                            _RPT.Fields.Item("PymntGroup").Value.ToString());
                    _RPT.MoveNext();
                }
                cmbPayT.Item.DisplayDesc = true;

                // === Account Payable ===
                SAPbouiCOM.ComboBox cmbAcct = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbAcct").Specific;
                ClearValidValues(cmbAcct.ValidValues);

                string accPayable = "SELECT DISTINCT \"AcctCode\", \"AcctName\" FROM OACT WHERE \"LocManTran\" = 'Y'";
                Utilities.ExecuteSQL(ref _RPTA, accPayable);
                while (!_RPTA.EoF)
                {
                    cmbAcct.ValidValues.Add(_RPTA.Fields.Item("AcctCode").Value.ToString(),
                                            _RPTA.Fields.Item("AcctName").Value.ToString());
                    _RPTA.MoveNext();
                }
                cmbAcct.Item.DisplayDesc = true;

                // === Bank Name Combo ===
                SAPbouiCOM.ComboBox cmbBank = (SAPbouiCOM.ComboBox)_Form.Items.Item("txtBankN").Specific;
                ClearValidValues(cmbBank.ValidValues);

                string bankName = "SELECT DISTINCT \"BankCode\", \"BankName\" FROM ODSC";
                Utilities.ExecuteSQL(ref _RPBN, bankName);
                while (!_RPBN.EoF)
                {
                    try
                    {
                        cmbBank.ValidValues.Add(_RPBN.Fields.Item("BankCode").Value.ToString(),
                                                _RPBN.Fields.Item("BankName").Value.ToString());
                    }
                    catch {  }
                    _RPBN.MoveNext();
                }
                cmbBank.Item.DisplayDesc = true;

                _Form.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                _Form.Items.Item("txtStatus").Enabled = false;
                _Form.Items.Item("txtCrtBy").Enabled = false;
                _Form.Items.Item("txtCrtOn").Enabled = false;
                _Form.Items.Item("txtAppBy").Enabled = false;
                _Form.Items.Item("txtAppOn").Enabled = false;
                _Form.PaneLevel = 1;

                if (isFinance)
                {
                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    return;
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
                _Form.Visible = true;
            }
        }




        #region Add Sequence No
        private void AddSequenceNumbersToMatrix(Matrix oMatrix)
        {
            //oMatrix.Columns.Item("V_-1").Editable = true;
            oMatrix.Columns.Item("rowHead").Visible = true;
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                oMatrix.Columns.Item("rowHead").Cells.Item(i).Specific.Value = i.ToString();
            }
        }
        #endregion

        //#region Set New line
        //private void SetNewLine(string FormUID, int Row)
        //{
        //    _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("25").Specific;
        //    _ChildDBDataSource.Offset = Row - 1;
        //    AddSequenceNumbersToMatrix(_Matrix);
        //    _ChildDBDataSource.SetValue("U_Type", _ChildDBDataSource.Offset, Row.ToString());
        //    _ChildDBDataSource.SetValue("U_ItemCode", _ChildDBDataSource.Offset, "");
        //    _ChildDBDataSource.SetValue("U_ItemName", _ChildDBDataSource.Offset, "");
        //    _ChildDBDataSource.SetValue("U_Quantity", _ChildDBDataSource.Offset, "");
        //    _ChildDBDataSource.SetValue("U_UOM", _ChildDBDataSource.Offset, "");
        //    _ChildDBDataSource.SetValue("U_Whse", _ChildDBDataSource.Offset, "");
        //    _Matrix.SetLineData(Row);
        //    _Matrix.FlushToDataSource();
        //    Utilities.SerializedMartix(ref _Form, ref _Matrix);
        //}
        //#endregion

        //#region FILTER Item Type
        //internal void FilterItemType(string CFLId)
        //{

        //    _Form.Freeze(true);
        //    try
        //    {
        //        _Cons = new Conditions();
        //        _Con = _Cons.Add();
        //        _Con.BracketOpenNum = 1;
        //        _Con.Alias = "U_BOMItem";
        //        _Con.Operation = BoConditionOperation.co_EQUAL;
        //        _Con.CondVal = "YES";
        //        _Con.BracketCloseNum = 1;
        //        _Form.ChooseFromLists.Item("CFL_Itm").SetConditions(_Cons);
        //    }


        //    catch (Exception ex)
        //    {
        //        Utilities.ShowErrorMessage(ex.Message);
        //    }
        //    finally
        //    {
        //        _Form.Freeze(false);
        //    }

        //}

        //internal void FilterItemTypeSp(string CFLId)
        //{

        //    _Form.Freeze(true);
        //    try
        //    {
        //        _Cons = new Conditions();
        //        _Con = _Cons.Add();
        //        _Con.BracketOpenNum = 1;
        //        _Con.Alias = "InvntItem";
        //        _Con.Operation = BoConditionOperation.co_EQUAL;
        //        _Con.CondVal = "Y";
        //        _Con.BracketCloseNum = 1;
        //        _Form.ChooseFromLists.Item("CFL_Itm").SetConditions(_Cons);
        //    }


        //    catch (Exception ex)
        //    {
        //        Utilities.ShowErrorMessage(ex.Message);
        //    }
        //    finally
        //    {
        //        _Form.Freeze(false);
        //    }

        //}


        //internal void FilterItemLineType(string CFLId)
        //{

        //    _Form.Freeze(true);
        //    try
        //    {
        //        _Cons = new Conditions();

        //        _Con = _Cons.Add();
        //        _Con.BracketOpenNum = 1;
        //        _Con.Alias = "ItemType";
        //        _Con.Operation = BoConditionOperation.co_NOT_EQUAL;
        //        _Con.CondVal = "F";
        //        _Con.BracketCloseNum = 1;
        //        _Form.ChooseFromLists.Item("CFL_LItm").SetConditions(_Cons);

        //    }


        //    catch (Exception ex)
        //    {
        //        Utilities.ShowErrorMessage(ex.Message);
        //    }
        //    finally
        //    {
        //        _Form.Freeze(false);
        //    }

        //}
        //#endregion

        #region LOAD COMPANY SETTINGS
        public static DataSet LoadCompanySetting()
        {
            XmlDataDocument xmlDoc = new XmlDataDocument();
            DataSet dataSet = new DataSet();
            dataSet.ReadXml("AppConfig.xml");
            //dataSet.ReadXml("App.Config");
            return dataSet;

        }
        #endregion

        #region Add New Row
        private void AddNewRow()
        {
            SAPbouiCOM.DBDataSource dbData = default(SAPbouiCOM.DBDataSource);
            dbData = (SAPbouiCOM.DBDataSource)_Form.DataSources.DBDataSources.Item(1);
            int selRow = _Matrix.GetNextSelectedRow();
            dbData.Clear();
            _Matrix.AddRow(1, selRow - 1);
            AddSequenceNumbersToMatrix(_Matrix);
        }
        #endregion

        //#region Create BOM
        //private void AddUpdateBOM()
        //{
        //    if (((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbStatus").Specific).Value == "Active")
        //    {
        //        try
        //        {
        //            _PRNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value;
        //            _PRName = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNam").Specific).Value;
        //            _Qty = Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("txtQty").Specific).Value);
        //            _Warehouse = ((SAPbouiCOM.EditText)_Form.Items.Item("txtWhse").Specific).Value;
        //            _RevNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPPNo").Specific).Value;
        //            _RevRmrks = ((SAPbouiCOM.EditText)_Form.Items.Item("txtVRmrks").Specific).Value;
        //            int _PriceList = Convert.ToInt32(((SAPbouiCOM.EditText)_Form.Items.Item("txtPNo").Specific).Value);
        //            Utilities.ExecuteSQL(ref _RSBom, "Select T0.\"Code\" " +
        //               "From \"OITT\" T0   Where  T0.\"Code\" = '" + _PRNum + "'");
        //            // Create a BOM object
        //            SAPbobsCOM.ProductTrees _BOM = (SAPbobsCOM.ProductTrees)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees);
        //            StageID = 0;
        //            if (_Matrix.RowCount > 0)
        //            {
        //                if (_RSBom.RecordCount > 0)
        //                {
        //                    _BOM.TreeCode = _PRNum;
        //                    _BOM.ProductDescription = _PRName;
        //                    _BOM.Quantity = _Qty;
        //                    _BOM.Warehouse = _Warehouse;
        //                    _BOM.TreeType = SAPbobsCOM.BoItemTreeTypes.iProductionTree;
        //                    _BOM.UserFields.Fields.Item("U_PPNum").Value = _RevNum;
        //                    _BOM.UserFields.Fields.Item("U_Remarks").Value = _RevRmrks;
        //                    _BOM.PriceList = _PriceList;

        //                    for (int i = 1; i <= _Matrix.RowCount; i++)
        //                    {
        //                        if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value != "")
        //                        {
        //                            if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Route Stage")
        //                            {
        //                                StageID = (StageID + 1);
        //                                _BOM.Stages.StageEntry = StageID;
        //                                _BOM.Stages.Name = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_7").Cells.Item(i).Specific).Value;
        //                                _BOM.Stages.SequenceNumber = StageID;
        //                                _BOM.Stages.Add();
        //                            }
        //                            else if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Resource" || ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Item")
        //                            {
        //                                if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Item")
        //                                {
        //                                    _BOM.Items.ItemType = ProductionItemType.pit_Item;

        //                                }
        //                                else if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Resource")
        //                                {
        //                                    _BOM.Items.ItemType = ProductionItemType.pit_Resource;
        //                                }
        //                                _BOM.Items.ItemCode = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value;
        //                                _BOM.Items.Quantity = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_6").Cells.Item(i).Specific).Value);
        //                                _BOM.Items.Warehouse = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(i).Specific).Value;
        //                                if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_9").Cells.Item(i).Specific).Value == "BackFlush")
        //                                {
        //                                    _BOM.Items.IssueMethod = BoIssueMethod.im_Backflush;
        //                                }
        //                                else
        //                                {
        //                                    _BOM.Items.IssueMethod = BoIssueMethod.im_Manual;
        //                                }
        //                                _BOM.Items.PriceList = _PriceList;
        //                                _BOM.Items.StageID = StageID;
        //                                _BOM.Items.Add();
        //                            }


        //                        }
        //                    }
        //                    int Result = _BOM.Update();
        //                    if (Result == 0)
        //                    {
        //                        Utilities.ShowWarningMessage("Bills of Materials Updated successfully.");
        //                        errorFlag = false;
        //                    }
        //                    else
        //                    {
        //                        Utilities.Application.Company.GetLastError(out ErrCode, out ErrMsg);
        //                        Utilities.ShowErrorMessage(ErrMsg);
        //                        errorFlag = true;
        //                        return;
        //                    }

        //                }

        //                else
        //                {
        //                    _BOM.TreeCode = _PRNum;
        //                    _BOM.ProductDescription = _PRName;
        //                    _BOM.Quantity = _Qty;
        //                    _BOM.Warehouse = _Warehouse;
        //                    _BOM.TreeType = SAPbobsCOM.BoItemTreeTypes.iProductionTree;
        //                    _BOM.UserFields.Fields.Item("U_PPNum").Value = _RevNum;
        //                    _BOM.UserFields.Fields.Item("U_Remarks").Value = _RevRmrks;
        //                    _BOM.PriceList = _PriceList;
        //                    for (int i = 1; i <= _Matrix.RowCount; i++)
        //                    {
        //                        if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value != "")
        //                        {
        //                            if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Route Stage")
        //                            {
        //                                StageID = (StageID + 1);
        //                                _BOM.Stages.StageEntry = StageID;
        //                                _BOM.Stages.SequenceNumber = StageID;
        //                                _BOM.Stages.Name = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_7").Cells.Item(i).Specific).Value;
        //                                _BOM.Stages.Add();

        //                            }

        //                            else if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Resource" || ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Item")
        //                            {
        //                                if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Item")
        //                                {
        //                                    _BOM.Items.ItemType = ProductionItemType.pit_Item;

        //                                }
        //                                else if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Resource")
        //                                {
        //                                    _BOM.Items.ItemType = ProductionItemType.pit_Resource;
        //                                }
        //                                _BOM.Items.ItemCode = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value;
        //                                _BOM.Items.Quantity = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_6").Cells.Item(i).Specific).Value);
        //                                _BOM.Items.Warehouse = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(i).Specific).Value;
        //                                if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_9").Cells.Item(i).Specific).Value == "BackFlush")
        //                                {
        //                                    _BOM.Items.IssueMethod = BoIssueMethod.im_Backflush;
        //                                }
        //                                else
        //                                {
        //                                    _BOM.Items.IssueMethod = BoIssueMethod.im_Manual;
        //                                }
        //                                _BOM.Items.StageID = StageID;
        //                                _BOM.Items.PriceList = _PriceList;
        //                                _BOM.Items.Add();
        //                            }


        //                        }
        //                    }
        //                    int Result = _BOM.Add();
        //                    if (Result == 0)
        //                    {
        //                        Utilities.ShowWarningMessage("Bills of Materials Created successfully.");
        //                        errorFlag = false;
        //                    }
        //                    else
        //                    {
        //                        Utilities.Application.Company.GetLastError(out ErrCode, out ErrMsg);
        //                        Utilities.ShowErrorMessage(ErrMsg);
        //                        errorFlag = true;
        //                        return;
        //                    }

        //                }
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine($"Error: {ex.Message}");
        //            errorFlag = true;
        //            return;
        //        }
        //        finally
        //        {
        //            // Disconnect from the SAP Business One company database
        //            //Utilities.Application.Company.Disconnect();
        //        }



        //    }
        //}
        //#endregion

        //#region Create or Update Production Order
        //private void AddUpdateProductionOrder()
        //{
        //    try
        //    {
        //        string type = ((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbType").Specific).Value;
        //        string docNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPPNo").Specific).Value;
        //        string itemCode = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value;
        //        string warehouse = ((SAPbouiCOM.EditText)_Form.Items.Item("txtWhse").Specific).Value;
        //        double plannedQty = Convert.ToDouble(((SAPbouiCOM.EditText)_Form.Items.Item("txtQty").Specific).Value);
        //        DateTime dueDate = DateTime.ParseExact(((SAPbouiCOM.EditText)_Form.Items.Item("txtVDate").Specific).Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
        //        string docEntry = _Form.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0);
        //        // Check if production order already exists
        //        Utilities.ExecuteSQL(ref _RSBom, $"SELECT \"DocEntry\" FROM \"OWOR\" WHERE \"ItemCode\" = '{itemCode}' AND \"U_RevNum\" = '{docEntry}'");
        //        bool isUpdate = _RSBom.RecordCount > 0;
        //        SAPbobsCOM.ProductionOrders prodOrder = (SAPbobsCOM.ProductionOrders)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);
        //        if (!isUpdate)
        //        {
        //            prodOrder.ItemNo = itemCode;
        //            if(type=="Standard")
        //            {
        //                prodOrder.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard;
        //            }
        //           else if(type == "Special")
        //            {
        //                prodOrder.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial;
        //            }
        //            else
        //            {
        //                prodOrder.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotDisassembly;
        //            }
        //            prodOrder.PlannedQuantity = plannedQty;
        //            prodOrder.PostingDate = DateTime.Now;
        //            prodOrder.DueDate = dueDate;
        //            prodOrder.Warehouse = warehouse;
        //            prodOrder.UserFields.Fields.Item("U_RevNum").Value = docEntry; 
        //            prodOrder.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned;

        //            for (int i = 1; i <= _Matrix.RowCount; i++)
        //            {
        //                string lineItem = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value;
        //                if (!string.IsNullOrEmpty(lineItem))
        //                {
        //                    prodOrder.Lines.ItemType = SAPbobsCOM.ProductionItemType.pit_Item;
        //                    prodOrder.Lines.ItemNo = lineItem;
        //                    prodOrder.Lines.PlannedQuantity = Convert.ToDouble(((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_12").Cells.Item(i).Specific).Value);
        //                    prodOrder.Lines.Warehouse = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(i).Specific).Value;
        //                    prodOrder.Lines.Add();
        //                }
        //            }

        //            int result = prodOrder.Add();

        //            if (result == 0)
        //            {
        //                int newDocEntry = Convert.ToInt32(Utilities.Application.Company.GetNewObjectKey());
        //                _Form.DataSources.DBDataSources.Item(0).SetValue("U_PRODEntry", 0, newDocEntry.ToString());
        //                if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
        //                _Form.Items.Item("1").Click(); 
        //                Utilities.ShowWarningMessage("Production Order Created successfully.");
        //                _Form.Items.Item("btnCrtPO").Enabled = false;
        //                errorFlag = false;
        //            }
        //            else
        //            {
        //                Utilities.Application.Company.GetLastError(out ErrCode, out ErrMsg);
        //                Utilities.ShowErrorMessage(ErrMsg);
        //                errorFlag = true;
        //            }
        //        }
        //        else
        //        {
        //            Utilities.ShowErrorMessage("Production Order already exists for this production planning document.");
        //            errorFlag = true;
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        Utilities.ShowErrorMessage($"Error: {ex.Message}");
        //        errorFlag = true;
        //    }
        //}
        //#endregion

        //#region EnableFromMenu
        //private void EnableFromMenu()
        //{
        //    _Form.EnableMenu("1288", true);
        //    _Form.EnableMenu("1289", true);
        //    _Form.EnableMenu("1290", true);
        //    _Form.EnableMenu("1291", true);
        //    _Form.EnableMenu("1292", true);
        //    _Form.EnableMenu("1287", true);
        //    _Form.EnableMenu("779", true);
        //}
        //#endregion

        #region Disable Form item on Load
            private void Disable_Form_Item(string status)
            {
            string currentUser = Utilities.Application.SBO_Application.Company.UserName;


            string sql = "SELECT T0.\"USER_CODE\" " +
                         "FROM \"OUSR\" T0 " +
                         "WHERE T0.\"U_FinanceA\" = 'Yes' AND T0.\"USER_CODE\" = '" + currentUser + "'";

            Utilities.ExecuteSQL(ref _RSUsr, sql);

            if (!_RSUsr.EoF)
            {
                hasFinanceAccess = true;
            }

            if (status == "Submitted" && hasFinanceAccess)
            {
                _Form.Items.Item("txtCredit").Enabled = true;
                _Form.Items.Item("txtCmmtL").Enabled = true;
                _Form.Items.Item("cmbPayT").Enabled = true;
                _Form.Items.Item("cmbAcct").Enabled = true;
                _Form.Items.Item("txtBankN").Enabled = true;
                _Form.Items.Item("txtAcNo").Enabled = true;
                _Form.Items.Item("txtIFSC").Enabled = true;
                _Form.Items.Item("txtBBranch").Enabled = true;

                _Form.Items.Item("btnSubmit").Visible = false;
                _Form.Items.Item("btnSubmit").Enabled = false;
                _Form.Items.Item("btnAproved").Visible = true;
                _Form.Items.Item("btnAproved").Enabled = false;

                _Form.Items.Item("txtBPCode").Enabled = false;
                _Form.Items.Item("txtBPName").Enabled = false;
                _Form.Items.Item("txtPanNo").Enabled = false;
                _Form.Items.Item("cmbType").Enabled = false;
                _Form.Items.Item("cmbPrice").Enabled = false;
                _Form.Items.Item("cmbGroup").Enabled = false;
                _Form.Items.Item("txtBillTo").Enabled = false;
                _Form.Items.Item("txtShipTo").Enabled = false;
                _Form.Items.Item("txtName").Enabled = false;
                _Form.Items.Item("txtEmail").Enabled = false;
                _Form.Items.Item("txtPhone").Enabled = false;
             
            }

            if (status == "Submitted" && !hasFinanceAccess)
            {
                _Form.Items.Item("txtBPCode").Enabled = false;
                _Form.Items.Item("txtBPName").Enabled = false;
                _Form.Items.Item("txtPanNo").Enabled = false;
                _Form.Items.Item("cmbType").Enabled = false;
                _Form.Items.Item("cmbPrice").Enabled = false;
                _Form.Items.Item("cmbGroup").Enabled = false;
                _Form.Items.Item("txtBillTo").Enabled = false;
                _Form.Items.Item("txtShipTo").Enabled = false;
                _Form.Items.Item("txtName").Enabled = false;
                _Form.Items.Item("txtEmail").Enabled = false;
                _Form.Items.Item("txtPhone").Enabled = false;
                ((SAPbouiCOM.Matrix)_Form.Items.Item("mtx_A").Specific).Columns.Item("Col_Rmrks").Editable = false;
                ((SAPbouiCOM.Matrix)_Form.Items.Item("mtx_A").Specific).Columns.Item("Col_Entry").Editable = false;
            }
         
            if (status == "Submitted" && hasFinanceAccess && IsFinanceTabValid() && Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                _Form.Items.Item("btnAproved").Visible = true;
                _Form.Items.Item("btnAproved").Enabled = true;
            }

            if (status == "Awaiting Approval" && Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                _Form.Items.Item("btnAproved").Enabled = false;
                _Form.Items.Item("btnSubmit").Visible = false;

                _Form.Items.Item("txtCredit").Enabled = false;
                _Form.Items.Item("txtCmmtL").Enabled = false;
                _Form.Items.Item("cmbPayT").Enabled = false;
                _Form.Items.Item("cmbAcct").Enabled = false;
                _Form.Items.Item("txtBankN").Enabled = false;
                _Form.Items.Item("txtAcNo").Enabled = false;
                _Form.Items.Item("txtIFSC").Enabled = false;
                _Form.Items.Item("txtBBranch").Enabled = false;

                _Form.Items.Item("txtBPCode").Enabled = false;
                _Form.Items.Item("txtBPName").Enabled = false;
                _Form.Items.Item("txtPanNo").Enabled = false;
                _Form.Items.Item("cmbType").Enabled = false;
                _Form.Items.Item("cmbPrice").Enabled = false;
                _Form.Items.Item("cmbGroup").Enabled = false;
                _Form.Items.Item("txtBillTo").Enabled = false;
                _Form.Items.Item("txtShipTo").Enabled = false;
                _Form.Items.Item("txtName").Enabled = false;
                _Form.Items.Item("txtEmail").Enabled = false;
                _Form.Items.Item("txtPhone").Enabled = false;
                ((SAPbouiCOM.Matrix)_Form.Items.Item("mtx_A").Specific).Columns.Item("Col_Rmrks").Editable = false;
                ((SAPbouiCOM.Matrix)_Form.Items.Item("mtx_A").Specific).Columns.Item("Col_Entry").Editable = false;
            }

            if (status == "Draft" && Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                _Form.Items.Item("btnSubmit").Visible = true;
                _Form.Items.Item("btnSubmit").Enabled = true;
                
            }

            if (status == "Approved" && Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                _Form.Items.Item("btnAproved").Visible = false;
                _Form.Items.Item("btnSubmit").Visible = false;

                _Form.Items.Item("txtCredit").Enabled = false;
                _Form.Items.Item("txtCmmtL").Enabled = false;
                _Form.Items.Item("cmbPayT").Enabled = false;
                _Form.Items.Item("cmbAcct").Enabled = false;
                _Form.Items.Item("txtBankN").Enabled = false;
                _Form.Items.Item("txtAcNo").Enabled = false;
                _Form.Items.Item("txtIFSC").Enabled = false;
                _Form.Items.Item("txtBBranch").Enabled = false;

                _Form.Items.Item("txtBPCode").Enabled = false;
                _Form.Items.Item("txtBPName").Enabled = false;
                _Form.Items.Item("txtPanNo").Enabled = false;
                _Form.Items.Item("cmbType").Enabled = false;
                _Form.Items.Item("cmbPrice").Enabled = false;
                _Form.Items.Item("cmbGroup").Enabled = false;
                _Form.Items.Item("txtBillTo").Enabled = false;
                _Form.Items.Item("txtShipTo").Enabled = false;
                _Form.Items.Item("txtName").Enabled = false;
                _Form.Items.Item("txtEmail").Enabled = false;
                _Form.Items.Item("txtPhone").Enabled = false;
                ((SAPbouiCOM.Matrix)_Form.Items.Item("mtx_A").Specific).Columns.Item("Col_Rmrks").Editable = false;
                ((SAPbouiCOM.Matrix)_Form.Items.Item("mtx_A").Specific).Columns.Item("Col_Entry").Editable = false;
                _Form.Items.Item("btnBrowse").Enabled = false;
                _Form.Items.Item("btnDisplay").Enabled = false;
                _Form.Items.Item("btnDelete").Enabled = false;

                _Form.Items.Item("txtStatus").Enabled = false;
                _Form.Items.Item("txtCrtBy").Enabled = false;
                _Form.Items.Item("txtCrtOn").Enabled = false;
                _Form.Items.Item("txtAppBy").Enabled = false;
                _Form.Items.Item("txtAppOn").Enabled = false;
            }

            string currentCode = ((SAPbouiCOM.EditText)_Form.Items.Item("txtBPCode").Specific).Value;

            string sqlRe = "SELECT T0.\"U_IsRejected\" " +
                         "FROM \"@ABPMASTERH\" T0 " +
                         "WHERE T0.\"Code\" = '" + currentCode + "'";

            Utilities.ExecuteSQL(ref _RSRej, sqlRe);
            string isRejected = "";
            if (_RSRej.RecordCount > 0)
            {
                isRejected = _RSRej.Fields.Item(0).Value.ToString();
            }

            // Now check both
            if (status == "Draft" && isRejected == "Yes")
            {
                _Form.Items.Item("txtBPCode").Enabled = false;
            }
           

        }
        #endregion

        #region Send Alert
        void SendInternalAlert(string docCode, List<string> approverUserCodes, string alertType)
        {
            //SAPbobsCOM.Messages alert = (SAPbobsCOM.Messages)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages);

            //alert.Subject = "Approval Required: UDO Document";
            //alert.MessageText = $"A new document (Code: {docCode}) has been submitted and requires your approval.";

            SAPbobsCOM.Messages alert = (SAPbobsCOM.Messages)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages);

            if (alertType == "Submit")
            {
                alert.Subject = "BP Creation Alert";
                alert.MessageText = $"BP (Code: {docCode}) creation alert needs finance update.";
            }
            else if (alertType == "Approve")
            {
                alert.Subject = "Approval Required: BP Creation";
                alert.MessageText = $"BP (Code: {docCode}) has been updated by Finance and requires your approval.";
            }


            for (int i = 0; i < approverUserCodes.Count; i++)
            {
                alert.Recipients.Add();
                alert.Recipients.SetCurrentLine(i);
                alert.Recipients.UserCode = approverUserCodes[i];
                //alert.Recipients.NameTo = approverUserCodes[i]; 
                alert.Recipients.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
            }

            int ret = alert.Add();
            if (ret != 0)
            {
                string errMsg;
                int errCode;
                Utilities.Application.Company.GetLastError(out errCode, out errMsg);
                Utilities.Application.SBO_Application.MessageBox("Error sending alert: " + errMsg);
            }
            else
            {
                //Utilities.Application.SBO_Application.MessageBox("Approval request sent to: " + string.Join(", ", approverUserCodes));
                if (alertType == "Submit")
                    Utilities.Application.SBO_Application.MessageBox("Update alert sent to: " + string.Join(", ", approverUserCodes));
                else if (alertType == "Approve")
                    Utilities.Application.SBO_Application.MessageBox("Approval request sent to: " + string.Join(", ", approverUserCodes));
            }
        }
        #endregion

        //#region ShowFileDialog
        //private string ShowFileDialog()
        //{
        //    string selectedFile = null;
        //    Thread t = new Thread(() =>
        //    {
        //        OpenFileDialog dialog = new OpenFileDialog();
        //        dialog.Title = "Select a file";
        //        dialog.Filter = "All files (*.*)|*.*";

        //        if (dialog.ShowDialog() == DialogResult.OK)
        //        {
        //            selectedFile = dialog.FileName;
        //        }
        //    });

        //    t.SetApartmentState(ApartmentState.STA);
        //    t.Start();
        //    t.Join();

        //    return selectedFile;

        //}
        //#endregion


        #region SaveAttachmentToSAP
        private AttachmentResult SaveAttachmentToSAP(string sourceFile)
        {
            AttachmentResult result = new AttachmentResult { AbsEntry = -1, DestPath = "" };

            try
            {
                 FileName = Path.GetFileNameWithoutExtension(sourceFile);
                 FileType = Path.GetExtension(sourceFile).Replace(".", "");
                string attachPath = GetSAPAttachmentPath();
                string destPath = Path.Combine(attachPath, Path.GetFileName(sourceFile));
                result.DestPath = destPath;

                // Copy the file to SAP Attachments folder
                File.Copy(sourceFile, destPath, true);

                // Create SAP Attachment object
                SAPbobsCOM.Attachments2 oAttach = null;
                try
                {
                    oAttach = (SAPbobsCOM.Attachments2)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.oAttachments2);

                    oAttach.Lines.SourcePath = attachPath;
                    oAttach.Lines.FileName = FileName;
                    oAttach.Lines.FileExtension = FileType;

                    if (oAttach.Add() == 0)
                    {
                        result.AbsEntry = int.Parse(Utilities.Application.Company.GetNewObjectKey());
                    }
                    else
                    {
                        Utilities.Application.Company.GetLastError(out int errCode, out string errMsg);
                        Utilities.Application.SBO_Application.StatusBar.SetText($"Attachment failed: {errMsg}",BoMessageTime.bmt_Short,BoStatusBarMessageType.smt_Error);
                    }
                }
                finally
                {
                    if (oAttach != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAttach);
                        oAttach = null;
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText($"Error in SaveAttachmentToSAP: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }

            return result;
        }
        #endregion

        #region GetSAPAttachmentPath
        private string GetSAPAttachmentPath()
        {
            SAPbobsCOM.Recordset rs = null;
            try
            {
                rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery("SELECT \"AttachPath\" FROM \"OADP\"");

                return rs.Fields.Item(0).Value.ToString();
            }
            finally
            {
                if (rs != null)
                {
                    Marshal.ReleaseComObject(rs);
                    rs = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        #endregion

        #region OpenAttachment
        private void OpenAttachment(int absEntry)
        {
            SAPbobsCOM.Attachments2 attach = null;
            try
            {
                attach = (SAPbobsCOM.Attachments2)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.oAttachments2);
                if (attach.GetByKey(absEntry))
                {
                    string fullPath = Path.Combine(
                        attach.Lines.SourcePath,
                        attach.Lines.FileName + "." + attach.Lines.FileExtension
                    );

                    if (File.Exists(fullPath))
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                        {
                            FileName = fullPath,
                            UseShellExecute = true
                        });
                    }
                    else
                    {
                        Utilities.Application.SBO_Application.StatusBar.SetText(
                            $"File not found: {fullPath}",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error
                        );
                    }
                }
            }
            finally
            {
                if (attach != null)
                {
                    Marshal.ReleaseComObject(attach);
                    attach = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        #endregion

        #region AttachmentResult
        public class AttachmentResult
        {
            public int AbsEntry { get; set; }
            public string DestPath { get; set; }
        }

        #endregion

        #region FinanceTabValidation
        private bool IsFinanceTabValid()
        {
            string credit = ((SAPbouiCOM.EditText)_Form.Items.Item("txtCredit").Specific).Value;
            if (string.IsNullOrWhiteSpace(credit) || Convert.ToDouble(credit) <= 0) return false;
            string commit = ((SAPbouiCOM.EditText)_Form.Items.Item("txtCmmtL").Specific).Value;
            if (string.IsNullOrWhiteSpace(commit) || Convert.ToDouble(commit) <= 0) return false;

            //if (string.IsNullOrWhiteSpace(((SAPbouiCOM.EditText)_Form.Items.Item("txtBBranch").Specific).Value)) return false;
            //if (string.IsNullOrWhiteSpace(((SAPbouiCOM.ComboBox)_Form.Items.Item("txtBankN").Specific).Value)) return false;
            //if (string.IsNullOrWhiteSpace(((SAPbouiCOM.EditText)_Form.Items.Item("txtAcNo").Specific).Value)) return false;
            //if (string.IsNullOrWhiteSpace(((SAPbouiCOM.EditText)_Form.Items.Item("txtIFSC").Specific).Value)) return false;
            if (string.IsNullOrWhiteSpace(((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbPayT").Specific).Value)) return false;
            if (string.IsNullOrWhiteSpace(((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbAcct").Specific).Value)) return false;

            return true;
        }
        #endregion

        #region Load Accounts based on BP Type
        public void LoadAccountsPayable()
        { 
            try
            {
                SAPbouiCOM.ComboBox cmbAcct = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbAcct").Specific;

                while (cmbAcct.ValidValues.Count > 0)
                {
                    cmbAcct.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }

                string bpCode = _MasterDBDataSource.GetValue("Code", 0).Trim();
                string selectedType = string.Empty;

                if (!string.IsNullOrEmpty(bpCode))
                {
                    Recordset rsType = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                    string qryType = $@"SELECT ""U_BPType"" FROM ""@ABPMASTERH"" WHERE ""Code"" = '{bpCode}'";

                    rsType.DoQuery(qryType);

                    if (!rsType.EoF)
                        selectedType = rsType.Fields.Item("U_BPType").Value.ToString(); // "C" or "S"
                }

                if (!string.IsNullOrEmpty(selectedType))
                {
                    string qryAcct = string.Empty;

                    if (selectedType == "S")
                    {
                        qryAcct = @"SELECT ""AcctCode"", ""AcctName"" FROM OACT WHERE ""LocManTran"" = 'Y' AND ""GroupMask"" = 2";
                    }
                    else if (selectedType == "C")
                    {
                        qryAcct = @"SELECT ""AcctCode"", ""AcctName"" FROM OACT WHERE ""LocManTran"" = 'Y' AND ""GroupMask"" = 1";
                    }

                    if (!string.IsNullOrEmpty(qryAcct))
                    {
                        Utilities.ExecuteSQL(ref _RPTA, qryAcct);

                        while (!_RPTA.EoF)
                        {
                            string acctCode = _RPTA.Fields.Item("AcctCode").Value.ToString();
                            string acctName = _RPTA.Fields.Item("AcctName").Value.ToString();

                            cmbAcct.ValidValues.Add(acctCode, acctName);
                            cmbAcct.Item.DisplayDesc = true;

                            _RPTA.MoveNext();
                        }
                    }
                }
                else
                {
                    cmbAcct.ValidValues.Add("", "Select BP Type First");
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Error while loading Accounts: " + ex.Message);
            }
         
        }
        #endregion

        #region CurrentUsersDepartment
        private string GetCurrentUserDepartment()
        {
            string userCode = Utilities.Application.SBO_Application.Company.UserName;
            string deptQuery = $@"
            SELECT T1.""Name"" 
            FROM OUSR T0 
            INNER JOIN OUDP T1 ON T0.""Department"" = T1.""Code"" 
            WHERE T0.""USER_CODE"" = '{userCode}'";

            Recordset rsDept = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            rsDept.DoQuery(deptQuery);

            if (!rsDept.EoF)
            {
                return rsDept.Fields.Item(0).Value.ToString();
            }

            return string.Empty; // no department found
        }
        #endregion

        #region EnableForAddMode
        public void EnableForAddMode()
        {
            _Form.Items.Item("txtBPCode").Enabled = true;
            _Form.Items.Item("txtBPName").Enabled = true;
            _Form.Items.Item("txtPanNo").Enabled = true;
            _Form.Items.Item("cmbType").Enabled = true;
            _Form.Items.Item("cmbPrice").Enabled = true;
            _Form.Items.Item("cmbGroup").Enabled = true;
            _Form.Items.Item("txtBillTo").Enabled = true;
            _Form.Items.Item("txtShipTo").Enabled = true;
            _Form.Items.Item("txtName").Enabled = true;
            _Form.Items.Item("txtEmail").Enabled = true;
            _Form.Items.Item("txtPhone").Enabled = true;
            ((SAPbouiCOM.Matrix)_Form.Items.Item("mtx_A").Specific).Columns.Item("Col_Rmrks").Editable = true;
            ((SAPbouiCOM.Matrix)_Form.Items.Item("mtx_A").Specific).Columns.Item("Col_Entry").Editable = true;
            _Form.Items.Item("btnBrowse").Enabled = true;
            _Form.Items.Item("btnDisplay").Enabled = true;
            _Form.Items.Item("btnDelete").Enabled = true;

            try
            {
                string departmentName = GetCurrentUserDepartment();

                //if (departmentName.Equals("ACCOUNTS", StringComparison.OrdinalIgnoreCase))
                if (departmentName.Equals("ACCOUNTS", StringComparison.OrdinalIgnoreCase) ||
                departmentName.Equals("ACCOUNT", StringComparison.OrdinalIgnoreCase) ||
                departmentName.Equals("FINANCE", StringComparison.OrdinalIgnoreCase) ||
                departmentName.Equals("ACCOUNTING", StringComparison.OrdinalIgnoreCase))
                {
                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    return;
                }
                _Form.Freeze(true);
                //_Form.DataBrowser.BrowseBy = "txtEntry";
                _Form.Settings.Enabled = true;
                UserObjectsMD oUserObjectMD = (UserObjectsMD)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
                _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("mtx_A").Specific;
                _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@ABPMASTERH");
                _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@ABPMASTERC");
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtCrtBy").Specific).Value = Utilities.Application.SBO_Application.Company.UserName;
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtCrtOn").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtStatus").Specific).Value = "Draft";

                //Price List
                SAPbouiCOM.ComboBox cmbPrice = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbPrice").Specific;

                // Clear old values
                ClearValidValues(cmbPrice.ValidValues);

                // Reload from query
                string priceList = "SELECT \"ListNum\",\"ListName\" FROM OPLN ORDER BY \"ListName\" ASC";
                Utilities.ExecuteSQL(ref _RSPL, priceList);

                while (!_RSPL.EoF)
                {
                    string listId = _RSPL.Fields.Item("ListNum").Value.ToString();
                    string listName = _RSPL.Fields.Item("ListName").Value.ToString();

                    cmbPrice.ValidValues.Add(listId, listName);

                    _RSPL.MoveNext();
                }

                cmbPrice.Item.DisplayDesc = true;



                //Header CardType
                SAPbouiCOM.ComboBox cmbType = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbType").Specific;
                ClearValidValues(cmbType.ValidValues);

                string qryCardType = "SELECT DISTINCT \"CardType\", " +
                                     "CASE WHEN \"CardType\" = 'C' THEN 'Customer' " +
                                     "     WHEN \"CardType\" = 'S' THEN 'Vendor' " +
                                     "END AS \"CardDesc\" FROM OCRD";

                Utilities.ExecuteSQL(ref _RSBPT, qryCardType);

                while (!_RSBPT.EoF)
                {
                    string cardType = _RSBPT.Fields.Item("CardType").Value.ToString();
                    string cardDesc = _RSBPT.Fields.Item("CardDesc").Value.ToString();

                    cmbType.ValidValues.Add(cardType, cardDesc);

                    _RSBPT.MoveNext();
                }
                cmbType.Item.DisplayDesc = true;


                //Header BPGroup
                SAPbouiCOM.ComboBox cmbGroup = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbGroup").Specific;
                ClearValidValues(cmbGroup.ValidValues);

                string group = "SELECT DISTINCT \"GroupCode\", \"GroupName\" FROM OCRG ORDER BY \"GroupName\" ASC";
                Utilities.ExecuteSQL(ref _RSPLG, group);

                while (!_RSPLG.EoF)
                {
                    string groupCode = _RSPLG.Fields.Item("GroupCode").Value.ToString();
                    string groupName = _RSPLG.Fields.Item("GroupName").Value.ToString();

                    cmbGroup.ValidValues.Add(groupCode, groupName);

                    _RSPLG.MoveNext();
                }
                cmbGroup.Item.DisplayDesc = true;



            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
                _Form.Visible = true;
            }
            _Form.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            //_Form.Items.Item("btnSubmit").Enabled = false;
            _Form.Items.Item("txtStatus").Enabled = false;
            _Form.Items.Item("txtCrtBy").Enabled = false;
            _Form.Items.Item("txtCrtOn").Enabled = false;
            _Form.Items.Item("txtAppBy").Enabled = false;
            _Form.Items.Item("txtAppOn").Enabled = false;

            _Form.Freeze(false);
            _Form.PaneLevel = 1;
        }
        #endregion

        #region clear valid values
        private void ClearValidValues(SAPbouiCOM.ValidValues validValues)
        {
            for (int i = validValues.Count - 1; i >= 0; i--)
            {
                validValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
            }
        }
        #endregion


        //#region Route line setup
        //private void Route_Line_Set(Matrix _Matrix)
        //{
        //    for (int i = 1; i <= _Matrix.RowCount; i++)
        //    {

        //        if (_Matrix.Columns.Item("V_8").Cells.Item(i).Specific.Value == "Route Stage")
        //        {
        //            _Matrix.CommonSetting.SetCellFontSize(i, 1, 10);
        //            _Matrix.CommonSetting.SetCellFontSize(i, 2, 10);
        //            _Matrix.CommonSetting.SetCellFontSize(i, 3, 10);
        //            _Matrix.CommonSetting.SetCellFontStyle(i, 1, BoFontStyle.fs_Bold);
        //            _Matrix.CommonSetting.SetCellFontStyle(i, 2, BoFontStyle.fs_Bold);
        //            _Matrix.CommonSetting.SetCellFontStyle(i, 3, BoFontStyle.fs_Bold);
        //            //SetCellForeColor(pVal.Row, 1, greenColor);
        //            _Matrix.CommonSetting.SetCellEditable(i, 4, false);
        //            _Matrix.CommonSetting.SetCellEditable(i, 6, false);
        //            _Matrix.CommonSetting.SetCellEditable(i, 7, false);
        //        }
        //    }

        //}
        //#endregion

        //#region Load Matrix Based on Type
        //private void LoadProductionPC(string orderType, string itemCode)
        //{
        //    _Form.Freeze(true);

        //    var childDS = _Form.DataSources.DBDataSources.Item("@PRODUCTIONPC");
        //    while (childDS.Size > 0)
        //        childDS.RemoveRecord(0);

        //    if (orderType == "Standard" || orderType == "Disassembly")
        //    {
        //        string query = $@"
        //        SELECT T1.""Code"", T2.""ItemName"", T1.""Quantity"", T2.""InvntryUom"",T0.""Qauntity"",T1.""Warehouse"",T1.""IssueMthd""
        //        FROM ""OITT"" T0 
        //        INNER JOIN ""ITT1"" T1 ON T0.""Code"" = T1.""Father""
        //        INNER JOIN ""OITM"" T2 ON T1.""Code"" = T2.""ItemCode""
        //        WHERE T0.""Code"" = '{itemCode.Replace("'", "''")}'";
        //        Utilities.ExecuteSQL(ref _RSBomP, query);
        //         while (!_RSBomP.EoF)
        //            {
        //                childDS.InsertRecord(childDS.Size);
        //                int dsIndex = childDS.Size - 1;

        //                childDS.SetValue("U_Type", dsIndex, "Item");
        //                childDS.SetValue("U_ItemCode", dsIndex, _RSBomP.Fields.Item("Code").Value.ToString().Trim());
        //                childDS.SetValue("U_ItemName", dsIndex, _RSBomP.Fields.Item("ItemName").Value.ToString().Trim());
        //                //double qty = Convert.ToDouble(_RSBomP.Fields.Item("Quantity").Value.ToString());
        //                childDS.SetValue("U_Quantity", dsIndex, _RSBomP.Fields.Item("Quantity").Value.ToString());
        //                childDS.SetValue("U_BaseQty", dsIndex, (_RSBomP.Fields.Item("Quantity").Value / _RSBomP.Fields.Item("Qauntity").Value));
        //                childDS.SetValue("U_PlannedQty", dsIndex, (_RSBomP.Fields.Item("Quantity").Value / _RSBomP.Fields.Item("Qauntity").Value));
        //                childDS.SetValue("U_UOM", dsIndex, _RSBomP.Fields.Item("InvntryUom").Value.ToString().Trim());
        //                childDS.SetValue("U_Whse", dsIndex, _RSBomP.Fields.Item("Warehouse").Value.ToString().Trim());
        //                if (_RSBomP.Fields.Item("IssueMthd").Value.ToString().Trim() == "B")
        //                {
        //                    childDS.SetValue("U_IssueMethod", dsIndex, "BackFlush");
        //                }
        //                else
        //                {
        //                    childDS.SetValue("U_IssueMethod", dsIndex, "Manual");
        //                }
        //                _RSBomP.MoveNext();
        //            }
        //        _Matrix.LoadFromDataSource();
        //    }
        //    else 
        //    {
        //        //_Form.Freeze(true);
        //        var childDSs = _Form.DataSources.DBDataSources.Item("@PRODUCTIONPC");
        //        while (childDSs.Size > 0)
        //            childDSs.RemoveRecord(0);
        //        childDSs.InsertRecord(0);
        //        _Matrix.LoadFromDataSource();
        //        _Matrix.AddRow(1, -1);
        //        //_Form.Freeze(false);
        //    }

        //    _Form.Freeze(false);

        //}

        //#endregion


    }
}
