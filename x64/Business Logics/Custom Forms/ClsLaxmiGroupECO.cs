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
    class ClsLaxmiGroupECO:Base
    {
        //#region VARIABLE DECLARATION

        //SAPbouiCOM.Item _ExistItem, _NewItem, _Items;
        //private bool flag;
        //private bool errorFlag;
        //public SAPbouiCOM.Matrix _Matrix;
        //public SAPbouiCOM.Grid _Grid;
        //SAPbouiCOM.StaticText _StaticText;
        //SAPbouiCOM.Button _Button;
        //string _DocEntry = "";
        //string _PRNum = string.Empty, _PRName = string.Empty, _RevNum = string.Empty, _Warehouse = string.Empty, _RevRmrks = string.Empty;
        //double _Qty = 0;
        //public string ErrMsg = string.Empty;
        //public int ErrCode = 0;
        //CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        //string ReportPath = string.Empty;
        //private static HanaConnection _conn;
        //private static HanaDataAdapter _da;
        ////private static DataTable _dataTable;
        //private static HanaParameter _parm1;
        //private static HanaParameter _parm2;
        //private static int _errors = 0;
        //string strcon, _Sql;
        //private SAPbouiCOM.EditText _EditText;
        //private SAPbouiCOM.LinkedButton _Linked;
        //private SAPbouiCOM.ComboBox _ComboBoxM;
        //private SAPbouiCOM.ComboBox _ComboBoxC, _ComboBoxCRow;
        //private SAPbouiCOM.ComboBox _ComboBoxB;
        //private SAPbouiCOM.ComboBox _ComboBoxL;
        //private SAPbouiCOM.ComboBox _ComboBoxH;
        //private SAPbouiCOM.ComboBox _ComboBox;
        //private SAPbouiCOM.DBDataSource _MasterDBDataSource = null, _ChildDBDataSource = null;
        //private SAPbouiCOM.ChooseFromListCollection _CFLS;
        //private SAPbouiCOM.ChooseFromList _CFL;
        //private SAPbouiCOM.ChooseFromListCreationParams _CFLC;

        //private SAPbouiCOM.Conditions _Cons;
        //private SAPbouiCOM.Condition _Con;
        //private int RowID { get; set; }
        //private string _Code { get; set; }
        //private string _LineId { get; set; }
        //private string _ActivityType = string.Empty;
        //string _ItemCode = string.Empty, _ItemName = string.Empty;
        //private SAPbobsCOM.Recordset _Rs, _RS, _rs, _RSD, _RSBom, _RSProd, _RSRoute;
        //SAPbouiCOM.ChooseFromListEvent _SysCFLEvent = null;
        //SAPbouiCOM.DBDataSource _DBData = null;
        //private System.Collections.Generic.List<string> ListOfBlockedID;
        //internal string _ServiceID { get; set; }
        //internal string _ActivityTypes { get; set; }
        //public bool chkStatus = true;


        //#endregion

        //#region CONSTRUCTOR & DISTRUCTOR
        //public ClsLaxmiGroupECO()
        //    : base()
        //{
        //    //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
        //    //
        //    // TODO: Add constructor logic here
        //    //
        //}
        //~ClsLaxmiGroupECO()
        //{
        //}
        //#endregion

        //#region ItemEvent
        //public override void Item_Event(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
        //{
        //    #region Before Action
        //    if (pVal.BeforeAction)
        //    {
        //        switch (pVal.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                       
        //                if (pVal.ItemUID == "5" && pVal.CharPressed != 9 && pVal.InnerEvent == false)
        //                {
        //                    oBubbleEvent = false;
        //                    return;
        //                }
        //                if (pVal.ItemUID == "6" && pVal.CharPressed != 9 && pVal.InnerEvent == false)
        //                {
        //                    oBubbleEvent = false;
        //                    return;
        //                }

        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //                if (pVal.ColUID == "V_0")
        //                {
        //                    _ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific;
        //                    if (_ComboBoxC.Selected.Value == "Resource")
        //                    {
        //                        //string ColVal = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific).Value.ToString();
        //                        SAPbouiCOM.Column _Column = (SAPbouiCOM.Column)_Matrix.Columns.Item("V_0");
        //                        SAPbouiCOM.LinkedButton _LinkedButton = ((SAPbouiCOM.LinkedButton)(_Column.ExtendedObject));
        //                        _LinkedButton.LinkedObjectType = "290";
        //                    }
        //                    else
        //                    {
        //                        // ColVal = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific).Value.ToString();
        //                        SAPbouiCOM.Column _Column = (SAPbouiCOM.Column)_Matrix.Columns.Item("V_0");
        //                        SAPbouiCOM.LinkedButton _LinkedButton = ((SAPbouiCOM.LinkedButton)(_Column.ExtendedObject));
        //                        _LinkedButton.LinkedObjectType = "4";
        //                    }
        //                }
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

        //                if (pVal.ItemUID == "1" && (Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
        //                {
        //                    _PRNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value;
        //                    _RevNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtRvnNo").Specific).Value;
        //                    _MasterDBDataSource.SetValue("U_RevNum", 0, (Utilities.getMaxColumnValueNumWithItem("@LGBOMH", "U_RevNum", _PRNum)));
        //                    if (((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value == "")
        //                    {
        //                        Utilities.Message("Product No can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                        oBubbleEvent = false;
        //                        return;
        //                    }
        //                    else if (((SAPbouiCOM.EditText)_Form.Items.Item("txtVDate").Specific).Value == "")
        //                    {
        //                        Utilities.Message("Valid Until date can not left blank. ", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                        oBubbleEvent = false;
        //                        return;
        //                    }
        //                    else if (((SAPbouiCOM.EditText)_Form.Items.Item("txtQty").Specific).Value == "")
        //                    {
        //                        Utilities.Message("Quantity can not left blank. ", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                        oBubbleEvent = false;
        //                        return;
        //                    }
        //                    else if (((SAPbouiCOM.EditText)_Form.Items.Item("txtWhse").Specific).Value == "")
        //                    {
        //                        Utilities.Message("Warehouse can not left blank. ", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                        oBubbleEvent = false;
        //                        return;
        //                    }
        //                    if (((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbStatus").Specific).Value == "")
        //                    {
        //                        Utilities.Message("Status can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                        oBubbleEvent = false;
        //                        return;
        //                    }
        //                    if (((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbPList").Specific).Value == "")
        //                    {
        //                        Utilities.Message("Price List can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                        oBubbleEvent = false;
        //                        return;
        //                    }
        //                    //if (Validation() == false)
        //                    //{
        //                    //    oBubbleEvent = false;
        //                    //    return;
        //                    //}
        //                    errorFlag = false;
        //                    //AddUpdateBOM();
        //                    //if (errorFlag == true)
        //                    //{
        //                    //    oBubbleEvent = false;
        //                    //    return;
        //                    //}
        //                }
        //                if (pVal.ItemUID == "1" && Form.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
        //                {
        //                    //Disable_Form_item();
        //                }
        //                if (pVal.ItemUID == "1" && (Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
        //                {
        //                    _PRNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value;
        //                    _RevNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtRvnNo").Specific).Value;
        //                    if (((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value == "")
        //                    {
        //                        Utilities.Message("Product No can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                        oBubbleEvent = false;
        //                        return;
        //                    }
        //                    else if (((SAPbouiCOM.EditText)_Form.Items.Item("txtVDate").Specific).Value == "")
        //                    {
        //                        Utilities.Message("Valid Until date can not left blank. ", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                        oBubbleEvent = false;
        //                        return;
        //                    }
        //                    else if (((SAPbouiCOM.EditText)_Form.Items.Item("txtSDate").Specific).Value == "")
        //                    {
        //                        Utilities.Message("Start date can not left blank. ", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                        oBubbleEvent = false;
        //                        return;
        //                    }
        //                    else if (((SAPbouiCOM.EditText)_Form.Items.Item("txtQty").Specific).Value == "")
        //                    {
        //                        Utilities.Message("Quantity can not left blank. ", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                        oBubbleEvent = false;
        //                        return;
        //                    }
        //                    else if (((SAPbouiCOM.EditText)_Form.Items.Item("txtWhse").Specific).Value == "")
        //                    {
        //                        Utilities.Message("Warehouse can not left blank. ", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                        oBubbleEvent = false;
        //                        return;
        //                    }
        //                    if (((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbStatus").Specific).Value == "")
        //                    {
        //                        Utilities.Message("Status can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                        oBubbleEvent = false;
        //                        return;
        //                    }
        //                    if (((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbPList").Specific).Value == "")
        //                    {
        //                        Utilities.Message("Price List can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                        oBubbleEvent = false;
        //                        return;
        //                    }
        //                    //if (Validation() == false)
        //                    //{
        //                    //    oBubbleEvent = false;
        //                    //    return;
        //                    //}
        //                    errorFlag = false;
        //                    //AddUpdateBOM();
        //                    if (errorFlag == true)
        //                    {
        //                        oBubbleEvent = false;
        //                        return;
        //                    }
        //                }
        //                if (pVal.ItemUID == "25")
        //                {
        //                    _Form.PaneLevel = 1;
        //                }
        //                else if (pVal.ItemUID == "26")
        //                {
        //                    _Form.PaneLevel = 2;
        //                }
        //                break;

        //            //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //            //    if (pVal.ItemUID == "5" && pVal.CharPressed != 9 && pVal.InnerEvent == false)
        //            //    {
        //            //        oBubbleEvent = false;
        //            //        return;
        //            //    }
        //            //    if (pVal.ItemUID == "6" && pVal.CharPressed != 9 && pVal.InnerEvent == false)
        //            //    {
        //            //        oBubbleEvent = false;
        //            //        return;
        //            //    }
        //            //    break;

        //            case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //                if (pVal.ItemUID == "txtPrNo")
        //                {
        //                    //this.FilterItemType("");
        //                }
        //                if (pVal.ItemUID == "27" && pVal.ColUID == "V_0")
        //                {
        //                    //this.FilterItemLineType("");
        //                }
        //                break;
        //        }
        //    }

        //    #endregion

        //    #region After Action
        //    else
        //    {
        //        switch (pVal.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //                //if (pVal.ItemUID == "27" && pVal.ColUID == "V_8" && pVal.Action_Success == true)
        //                //{
        //                //    bool chk = pVal.ActionSuccess;
        //                //    bool chk1 = pVal.Action_Success;
        //                //    bool chk2 = pVal.InnerEvent;
        //                //    _ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific;
        //                //    if (_ComboBoxC.Selected.Value == "Resource")
        //                //    {
        //                //        _CFLS = _Form.ChooseFromLists;
        //                //        _CFLC = Utilities.Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
        //                //        _CFLC.MultiSelection = true;
        //                //        _CFLC.ObjectType = "290";
        //                //       // _CFLC.UniqueID = pVal.Row.ToString();
        //                //        _CFL = _CFLS.Add(_CFLC);
        //                //        _Matrix.Columns.Item("V_0").ChooseFromListUID = pVal.Row.ToString();
        //                //        _Matrix.Columns.Item("V_0").ChooseFromListAlias = "VisResCode";

        //                //        string ColVal = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific).Value.ToString();
        //                //        SAPbouiCOM.Column _Column = (SAPbouiCOM.Column)_Matrix.Columns.Item("V_0");
        //                //        SAPbouiCOM.LinkedButton _LinkedButton = ((SAPbouiCOM.LinkedButton)(_Column.ExtendedObject));
        //                //        _LinkedButton.LinkedObjectType = "290";
        //                //    }
        //                //    else
        //                //    {
        //                //        _CFLS = _Form.ChooseFromLists;
        //                //        _CFLC = Utilities.Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
        //                //        _CFLC.MultiSelection = true;
        //                //        _CFLC.ObjectType = "4";
        //                //       // _CFLC.UniqueID.Replace("CFL_LItm", "");
        //                //       // _CFLC.UniqueID = pVal.Row.ToString();
        //                //        _CFL = _CFLS.Add(_CFLC);
        //                //        _Matrix.Columns.Item("V_0").ChooseFromListUID = pVal.Row.ToString();
        //                //        _Matrix.Columns.Item("V_0").ChooseFromListAlias = "ItemCode";

        //                //        string ColVal = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific).Value.ToString();
        //                //        SAPbouiCOM.Column _Column = (SAPbouiCOM.Column)_Matrix.Columns.Item("V_0");
        //                //        SAPbouiCOM.LinkedButton _LinkedButton = ((SAPbouiCOM.LinkedButton)(_Column.ExtendedObject));
        //                //        _LinkedButton.LinkedObjectType = "290";
        //                //    }

        //                //}

        //                //if (pVal.ItemUID == "7" && pVal.ColUID == "V_7")
        //                //{
        //                //    //_ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific;
        //                //    _ComboBoxC.Select(2, SAPbouiCOM.BoSearchKey.psk_Index);

        //                //}
        //                if (pVal.ItemUID == "cmbPList")
        //                {
        //                    _ComboBoxM = (SAPbouiCOM.ComboBox)Form.Items.Item("cmbPList").Specific;
        //                    _EditText = (SAPbouiCOM.EditText)Form.Items.Item("txtPNo").Specific;
        //                    _EditText.Value = _ComboBoxM.Selected.Description.ToString();
        //                }
        //                break;
        //                //if(pVal.ItemUID=="7"&& pVal.ColUID=="V_6")
        //                //{
        //                //    _Matrix.AddRow(1, pVal.Row);
        //                //}
        //                //if (pVal.ItemUID == "7" && pVal.ColUID == "V_2")
        //                //{
        //                //    _ComboBoxB = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific;
        //                //    _EditText = (SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific;
        //                //    _EditText.Value = _ComboBoxB.Selected.Description.ToString();
        //                //}
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //                if (pVal.ItemUID == "1000010")
        //                {
        //                    string _ItemCode = ((SAPbouiCOM.EditText)_Form.Items.Item("txtCItmCod").Specific).Value.ToString();
        //                    FilterData(_ItemCode);
        //                }
        //                if (pVal.ItemUID == "1" && Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //                {
        //                    Utilities.SerializedMartix(ref _Form, ref _Matrix);
        //                }
        //                //if (pVal.ItemUID == "1" && Form.Mode==SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //                //{
        //                //    AddUpdateBOM();
        //                //}
        //                //if (pVal.ItemUID == "27" && pVal.EventType == BoEventTypes.et_CLICK)
        //                //{
        //                //    if (pVal.ColUID == "V_0")
        //                //    {
        //                //        if (_ComboBoxC.Selected.Value == "Resource")
        //                //        {
        //                //            _Matrix.Columns.Item("V_0").ChooseFromListUID = "CFL_Rsc";
        //                //            _Matrix.Columns.Item("V_0").ChooseFromListAlias = "VisResCode";
        //                //        }
        //                //        else if (_ComboBoxC.Selected.Value == "Item")
        //                //        {
        //                //            _Matrix.Columns.Item("V_0").ChooseFromListUID = "CFL_LItm";
        //                //            _Matrix.Columns.Item("V_0").ChooseFromListAlias = "ItemCode";
        //                //        }
        //                //    }

        //                //}
        //                break;
        //            #region Set CFL Values
        //            case BoEventTypes.et_CHOOSE_FROM_LIST:
        //                _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
        //                if (pVal.ItemUID == "txtCItmCod" && _SysCFLEvent.SelectedObjects != null)
        //                {
        //                    _MasterDBDataSource.SetValue("U_CItmCod", 0, _SysCFLEvent.SelectedObjects.GetValue("ItemCode", 0).ToString().Trim());
        //                    _MasterDBDataSource.SetValue("U_CItmNam", 0, _SysCFLEvent.SelectedObjects.GetValue("ItemName", 0).ToString().Trim());
        //                    //_MasterDBDataSource.SetValue("U_CItmGrp", 0, _SysCFLEvent.SelectedObjects.GetValue("InvntryUom", 0).ToString().Trim());
        //                    //AddSequenceNumbersToMatrix(_Matrix);
        //                }

        //                if (pVal.ItemUID == "txtNItmCod" && _SysCFLEvent.SelectedObjects != null)
        //                {
        //                    _MasterDBDataSource.SetValue("U_NItmCod", 0, _SysCFLEvent.SelectedObjects.GetValue("ItemCode", 0).ToString().Trim());
        //                    _MasterDBDataSource.SetValue("U_NItmNam", 0, _SysCFLEvent.SelectedObjects.GetValue("ItemName", 0).ToString().Trim());
        //                    //_MasterDBDataSource.SetValue("U_NItmGrp", 0, _SysCFLEvent.SelectedObjects.GetValue("Invntry", 0).ToString().Trim());
        //                    //AddSequenceNumbersToMatrix(_Matrix);
        //                }
        //                break;
        //                #endregion


        //        }
        //    }
        //    #endregion
        //}
        //#endregion

        //#region FORM DEFAULT
        //internal void FormDefault()
        //{
        //    try
        //    {
        //        _Form.Freeze(true);
        //        _Form.EnableMenu("1292", true);
        //        _MasterDBDataSource = Form.DataSources.DBDataSources.Item("@LGECOH");
        //        _ChildDBDataSource = Form.DataSources.DBDataSources.Item("@LGECOC");
        //        _Form.DataBrowser.BrowseBy = "txtDocNum";
        //        ((SAPbouiCOM.EditText)_Form.Items.Item("txtDocNum").Specific).Value = Utilities.getMaxColumnValueNum("@LGECOH", "U_DocNo");
        //        ((SAPbouiCOM.EditText)_Form.Items.Item("txtDocDate").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
        //        ((SAPbouiCOM.EditText)_Form.Items.Item("txtPostBy").Specific).Value = Utilities.Application.SBO_Application.Company.UserName;
        //        string time = DateTime.Now.ToString("HHmmss");
        //        ((SAPbouiCOM.EditText)_Form.Items.Item("txtDocTime").Specific).Value = DateTime.Now.ToString("HHmm");
        //        //string sql =Utilities.Application.SBO_Application.Company.UserName;
        //        //((SAPbouiCOM.EditText)_Form.Items.Item("41").Specific).Active = true;
        //        _ComboBox = (SAPbouiCOM.ComboBox)_Form.Items.Item("cmbOpTyp").Specific;
        //        _ComboBox.ValidValues.Add("Replace", "Replace");
        //        _ComboBox.ValidValues.Add("Delete", "Delete");
        //        _ComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //        _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("mtxRslt").Specific;
        //        _Matrix.Columns.Item("V_0").Editable = false;
        //        _Matrix.Columns.Item("V_1").Editable = false;
        //        _Matrix.Columns.Item("V_2").Editable = false;
        //        _Matrix.Columns.Item("V_3").Editable = false;
        //        _Matrix.Columns.Item("V_4").Editable = false;
        //        _Matrix.Columns.Item("V_5").Editable = false;
        //        _Matrix.Columns.Item("V_6").Editable = false;
        //        _Matrix.Columns.Item("V_7").Editable = false;
        //        _Matrix.Columns.Item("V_8").Editable = false;
        //        _Matrix.Columns.Item("V_9").Editable = false;
        //        _Matrix.Columns.Item("V_10").Editable = false;
        //        _Matrix.Columns.Item("V_12").Editable = false;
        //        _Form.Items.Item("lblDocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //        _Form.Items.Item("txtDocNum").Enabled = false;
        //        _Form.Items.Item("txtDocDate").Enabled = false;
        //        _Form.Items.Item("txtDocTime").Enabled = false;
        //        _Form.Items.Item("txtPostBy").Enabled = false;
        //        _Form.Items.Item("txtDStatus").Enabled = false;
        //        //_Form.Items.Item("6").Enabled = false;
        //        //_Form.Items.Item("1000002").Enabled = false;
        //        //_Form.Items.Item("29").Enabled = false;
        //        //_Form.Items.Item("1000004").Enabled = false;
        //        //_Form.Items.Item("27").Enabled = false;
        //        //_Form.Items.Item("12").Enabled = false;
        //        //_Form.Items.Item("36").Enabled = false;
        //        //_Form.Items.Item("38").Enabled = false;
        //        //_Form.Items.Item("22").Enabled = false;
        //        //_Form.Items.Item("24").Enabled = false;

        //        Utilities.SerializedMartix(ref _Form, ref _Matrix);
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

        //}
        //#endregion

        //#region Filter Data
        //internal void FilterData(string _ItmCode)
        //{
        //    try
        //    {
        //        _Form.Freeze(true);
        //        string QueryFillData = string.Empty;

        //        _Matrix.Clear();
        //        Utilities.Application.SBO_Application.StatusBar.SetText("Please wait", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        //        _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("mtxRslt").Specific;

        //        QueryFillData = "select T0.\"U_PRNum\",T0.\"U_PRName\", T3.\"ItmsGrpNam\",T0.\"U_RevNum\",T0.\"U_RevRemarks\",T0.\"U_Status\", "+
        //                        "T1.\"LineId\" ,T1.\"U_Quantity\" from \"@LGBOMH\" T0 " +
        //                        "INNER JOIN \"@LGBOMC\" T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
        //                        "INNER JOIN OITM T2 on T2.\"ItemCode\" = T0.\"U_PRNum\" " +
        //                        "INNER JOIN OITB T3 on T3.\"ItmsGrpCod\" = T2.\"ItmsGrpCod\" " +
        //                        "where T1.\"U_ItemCode\" = '" + _ItmCode + "'";

        //        Utilities.ExecuteSQL(ref _RS, QueryFillData);

        //        Utilities.Application.SBO_Application.StatusBar.SetText("Please wait " + _RS.RecordCount + " lines data loading", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

        //        for (int i = 0; i < _RS.RecordCount; i++)
        //        {
        //            _Matrix.AddRow(1, -1);
        //            _Matrix.FlushToDataSource();
        //            _ChildDBDataSource.SetValue("U_PrdCod", i, _RS.Fields.Item("U_PRNum").Value.ToString());
        //            _ChildDBDataSource.SetValue("U_PrdNam", i, _RS.Fields.Item("U_PRNam").Value.ToString());
        //            _ChildDBDataSource.SetValue("U_PrdGrp", i, _RS.Fields.Item("ItmsGrpNam").Value.ToString());
        //            _ChildDBDataSource.SetValue("U_RevNum", i, _RS.Fields.Item("U_RevNum").Value.ToString());
        //            _ChildDBDataSource.SetValue("U_RevRmrks", i, _RS.Fields.Item("U_RevRemarks").Value.ToString());
        //            _ChildDBDataSource.SetValue("U_RevStatus", i, _RS.Fields.Item("U_Status").Value.ToString());
        //            _ChildDBDataSource.SetValue("U_LineNo", i, _RS.Fields.Item("LineId").Value.ToString());
        //            //_ChildDBDataSource.SetValue("U_RouteNo", i, _RS.Fields.Item("DocEntry").Value.ToString());
        //            //_ChildDBDataSource.SetValue("U_RouteName", i, _RS.Fields.Item("LineVat").Value.ToString());
        //            _ChildDBDataSource.SetValue("U_ExItmQty", i, _RS.Fields.Item("U_Quantity").Value.ToString());
        //            _Matrix.LoadFromDataSource();
        //            _RS.MoveNext();
        //        }

        //        _Form.Items.Item("txtCItmCod").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //        //_Form.Items.Item("10").Enabled = false;

        //        Utilities.SerializedMartix(ref _Form, ref _Matrix);
        //        Utilities.Message("Operation completed sucessfully.", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
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
        //}
        //#endregion
    }
}
