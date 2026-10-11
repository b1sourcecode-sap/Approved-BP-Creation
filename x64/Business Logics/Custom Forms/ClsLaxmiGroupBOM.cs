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
    class ClsLaxmiGroupBOM:Base
    {
        //#region VARIABLE DECLARATION

        //SAPbouiCOM.Item _ExistItem, _NewItem, _Items;
        //private bool flag;
        //private bool errorFlag;
        //public SAPbouiCOM.Matrix _Matrix;
        //SAPbouiCOM.StaticText _StaticText;
        //SAPbouiCOM.Button _Button;
        //string _DocEntry = "";
        //string _PRNum = string.Empty, _PRName = string.Empty, _RevNum =string.Empty,_Warehouse = string.Empty, _RevRmrks = string.Empty;
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
        //private SAPbobsCOM.Recordset _Rs, _RS, _rs, _RSD,_RSBom, _RSProd ,_RSRoute,_RSUsr,_RSMgItm;
        //SAPbouiCOM.ChooseFromListEvent _SysCFLEvent = null;
        //SAPbouiCOM.DBDataSource _DBData = null;
        //private System.Collections.Generic.List<string> ListOfBlockedID;
        //internal string _ServiceID { get; set; }
        //internal string _ActivityTypes { get; set; }
        //public bool chkStatus = true;
        //static int StageID = 0;
        //#endregion

        //#region CONSTRUCTOR & DISTRUCTOR
        //public ClsLaxmiGroupBOM()
        //    : base()
        //{
        //    //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
        //    //
        //    // TODO: Add constructor logic here
        //    //
        //}
        //~ClsLaxmiGroupBOM()
        //{
        //}
        //#endregion

        //#region FormDataEvent
        //public override void FormData_Event(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool oBubbleEvent)
        //{
        //    try
        //    {
        //        switch (BusinessObjectInfo.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //                Form.Freeze(true);
        //                Utilities.SerializedMartix(ref _Form, ref _Matrix);
        //                AddSequenceNumbersToMatrix(_Matrix);
        //                Route_Line_Set(_Matrix);
        //                Disable_Form_Item();
        //                Form.Freeze(false);
        //                break;
                       
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Utilities.ShowErrorMessage(ex.Message);
        //    }

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
        //                if (pVal.ColUID == "V_0"  && Form.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)//&& pVal.CharPressed == 9
        //                {
        //                    _ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific;
        //                    if(_ComboBoxC.Selected.Value == "Resource")
        //                    {
        //                        _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
        //                        DataSet ds = LoadCompanySetting();
        //                        string ServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
        //                        string UserID = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
        //                        string Password = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();
        //                        string CurrentSchema = ds.Tables[0].Rows[0]["Database"].ToString().Trim();
        //                        _conn = new HanaConnection("Server=" + ServerName + ";UserID=" + UserID + ";Password=" + Password + ";Current Schema=" + CurrentSchema + "");
        //                        this._Object = new ClsChooseFromList();
        //                        ((ClsChooseFromList)this._Object).FormTitle = "Resource List";
        //                        ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref oBubbleEvent, "select \"VisResCode\",\"ResName\",\"UnitOfMsr\" from \"ORSC\" ", strcon, true);
                               
        //                    }
        //                    else if (_ComboBoxC.Selected.Value == "Route Stage")
        //                    {
        //                        _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
        //                        DataSet ds = LoadCompanySetting();
        //                        string ServerName = ds.Tables[0].Rows[0]["DatabaseServerName"].ToString().Trim();
        //                        string UserID = ds.Tables[0].Rows[0]["dbUser"].ToString().Trim();
        //                        string Password = ds.Tables[0].Rows[0]["dbPassword"].ToString().Trim();
        //                        string CurrentSchema = ds.Tables[0].Rows[0]["Database"].ToString().Trim();
        //                        _conn = new HanaConnection("Server=" + ServerName + ";UserID=" + UserID + ";Password=" + Password + ";Current Schema=" + CurrentSchema + "");
        //                        this._Object = new ClsChooseFromList();
        //                        ((ClsChooseFromList)this._Object).FormTitle = "Route List";
        //                        ((ClsChooseFromList)this._Object).OpenCFL(this, pVal, ref oBubbleEvent, "select \"Code\",\"Desc\" from \"ORST\" ", strcon, true);
        //                    }
        //                }
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
        //                    else if(_ComboBoxC.Selected.Value == "Item")
        //                    {
        //                        //string ColVal = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific).Value.ToString();
        //                        SAPbouiCOM.Column _Column = (SAPbouiCOM.Column)_Matrix.Columns.Item("V_0");
        //                        SAPbouiCOM.LinkedButton _LinkedButton = ((SAPbouiCOM.LinkedButton)(_Column.ExtendedObject));
        //                        _LinkedButton.LinkedObjectType = "4";
        //                    }
        //                    else
        //                    {
        //                        //string ColVal = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific).Value.ToString();
        //                        SAPbouiCOM.Column _Column = (SAPbouiCOM.Column)_Matrix.Columns.Item("V_0");
        //                        SAPbouiCOM.LinkedButton _LinkedButton = ((SAPbouiCOM.LinkedButton)(_Column.ExtendedObject));
        //                        _LinkedButton.LinkedObjectType = "296";

        //                    }
        //                }
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                       
        //                if (pVal.ItemUID == "1" &&( Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
        //                {
        //                   _PRNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value;
        //                   _RevNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtRvnNo").Specific).Value;
        //                   _MasterDBDataSource.SetValue("U_RevNum", 0, (Utilities.getMaxColumnValueNumWithItem("@LGBOMH", "U_RevNum", _PRNum)));
        //                    if (((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value == "")
        //                    {
        //                        Utilities.Message("Product No can not left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                      oBubbleEvent = false;
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
        //                    if (Validation() == false)
        //                    {
        //                        oBubbleEvent = false;
        //                        return;
        //                    }
        //                    errorFlag = false;
        //                    AddUpdateBOM();
        //                    if(errorFlag== true)
        //                    {
        //                        oBubbleEvent = false;
        //                        return;
        //                    }
        //                }
        //                if (pVal.ItemUID == "1" && Form.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
        //                {
        //                    //Disable_Form_item();
        //                }
        //                if (pVal.ItemUID == "1" && (Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
        //                {
        //                    _PRNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value;
        //                    _RevNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtRvnNo").Specific).Value;
        //                    Utilities.ExecuteSQL(ref _RSProd, "Select T0.\"ItemCode\" " +
        //                   "From \"OWOR\" T0   Where T0.\"U_RevNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtRvnNo").Specific).Value + "' " +
        //                   "AND T0.\"ItemCode\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value + "'");

        //                    if (_RSProd.RecordCount>0 && ((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbStatus").Specific).Value=="Draft")
        //                    {
        //                        Utilities.Message("This Revision Alredy in Use, Can Not Change Status To Draft.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                        oBubbleEvent = false;
        //                        return;
        //                    }
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
        //                    if (Validation() == false)
        //                    {
        //                        oBubbleEvent = false;
        //                        return;
        //                    }
        //                    errorFlag = false;
        //                    AddUpdateBOM();
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
        //               break;
                       
        //            case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //                if (pVal.ItemUID == "txtPrNo")
        //                {
        //                    this.FilterItemType("");
        //                }
        //                if(pVal.ItemUID == "27" && pVal.ColUID == "V_0")
        //                {
        //                    this.FilterItemLineType("");
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

        //                if (pVal.ColUID == "V_8" && (Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
        //                {
        //                    _ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific;
        //                    if (_ComboBoxC.Selected.Value == "Route Stage")
        //                    {
        //                        _Matrix.CommonSetting.SetCellFontSize(pVal.Row, 1, 12);
        //                        _Matrix.CommonSetting.SetCellFontSize(pVal.Row, 2, 12);
        //                        _Matrix.CommonSetting.SetCellFontSize(pVal.Row, 3, 12);
        //                        _Matrix.CommonSetting.SetCellFontStyle(pVal.Row, 1, BoFontStyle.fs_Bold);
        //                        _Matrix.CommonSetting.SetCellFontStyle(pVal.Row, 2, BoFontStyle.fs_Bold);
        //                        _Matrix.CommonSetting.SetCellFontStyle(pVal.Row, 3, BoFontStyle.fs_Bold);
        //                        //SetCellForeColor(pVal.Row, 1, greenColor);
        //                        _Matrix.CommonSetting.SetCellEditable(pVal.Row, 4, false);
        //                        _Matrix.CommonSetting.SetCellEditable(pVal.Row, 6, false);
        //                        _Matrix.CommonSetting.SetCellEditable(pVal.Row, 7, false);
        //                    }

        //                }
        //                if (pVal.ItemUID == "cmbPList")
        //                {
        //                    _ComboBoxM = (SAPbouiCOM.ComboBox)Form.Items.Item("cmbPList").Specific;
        //                    _EditText = (SAPbouiCOM.EditText)Form.Items.Item("txtPNo").Specific;
        //                    _EditText.Value = _ComboBoxM.Selected.Description.ToString();
        //                }

        //                //if (pVal.ItemUID == "27" && pVal.ColUID == "V_8" && pVal.Action_Success == true)
        //                //{
        //                //    bool chk = pVal.ActionSuccess;
        //                //    bool chk1 = pVal.Action_Success;
        //                //    bool chk2 = pVal.InnerEvent;
        //                //    _ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific;
        //                //    if (_ComboBoxC.Selected.Value == "Item")
        //                //    {
        //                //        //try
        //                //        //{
        //                //            ChooseFromListCreationParams cflCreationParams = (ChooseFromListCreationParams)Utilities.Application.SBO_Application.CreateObject(BoCreatableObjectType.cot_ChooseFromListCreationParams);
        //                //            cflCreationParams.MultiSelection = false;
        //                //            cflCreationParams.ObjectType = "4"; // Business Partners
        //                //            cflCreationParams.UniqueID = pVal.Row.ToString();
        //                //            SAPbouiCOM.ChooseFromList oCFL = _Form.ChooseFromLists.Add(cflCreationParams);
        //                //            // Assign CFL to the Grid column
        //                //            SAPbouiCOM.Column _Column = (SAPbouiCOM.Column)_Matrix.Columns.Item("V_0");
        //                //            _Column.ChooseFromListUID = pVal.Row.ToString();
        //                //            _Column.ChooseFromListAlias = "ItemCode";
        //                //            SAPbouiCOM.LinkedButton _LinkedButton = ((SAPbouiCOM.LinkedButton)(_Column.ExtendedObject));
        //                //            _LinkedButton.LinkedObjectType = "4";
        //                //        //}
        //                //        //catch (Exception ex)
        //                //        //{

        //                //        //}

        //                //    }


        //                //}

        //                break;

        //            //if (pVal.ItemUID == "7" && pVal.ColUID == "V_7")
        //            //{
        //            //    //_ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific;
        //            //    _ComboBoxC.Select(2, SAPbouiCOM.BoSearchKey.psk_Index);

        //            //}

        //            //if(pVal.ItemUID=="7"&& pVal.ColUID=="V_6")
        //            //{
        //            //    _Matrix.AddRow(1, pVal.Row);
        //            //}
        //            //if (pVal.ItemUID == "7" && pVal.ColUID == "V_2")
        //            //{
        //            //    _ComboBoxB = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific;
        //            //    _EditText = (SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific;
        //            //    _EditText.Value = _ComboBoxB.Selected.Description.ToString();
        //            //}
        //            //break;
        //            case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //                //if (pVal.ItemUID == "1" && UpdateFlag== true)
        //                //{
        //                //    AddUpdateBOM();
        //                //}
        //                if (pVal.ItemUID == "1" && Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //                {
        //                    Utilities.SerializedMartix(ref _Form, ref _Matrix);
        //                }
        //                if (pVal.ItemUID == "1" && Form.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
        //                {
        //                    //Disable_Form_item();
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
        //                if (pVal.ItemUID == "txtPrNo" && _SysCFLEvent.SelectedObjects != null)
        //                {
        //                    _MasterDBDataSource.SetValue("U_PRNum", 0, _SysCFLEvent.SelectedObjects.GetValue("ItemCode", 0).ToString().Trim());
        //                    _MasterDBDataSource.SetValue("U_PRName", 0, _SysCFLEvent.SelectedObjects.GetValue("ItemName", 0).ToString().Trim());
        //                    _MasterDBDataSource.SetValue("U_UOM", 0, _SysCFLEvent.SelectedObjects.GetValue("InvntryUom", 0).ToString().Trim());
        //                    _MasterDBDataSource.SetValue("U_RevNum", 0, (Utilities.getMaxColumnValueNumWithItem("@LGBOMH", "U_RevNum", _SysCFLEvent.SelectedObjects.GetValue("ItemCode", 0).ToString().Trim())));
        //                    _MasterDBDataSource.SetValue("U_Qty", 0, "1");
        //                    _Matrix.AddRow(1, -1);
        //                    AddSequenceNumbersToMatrix(_Matrix);
        //                }
        //                if (pVal.ItemUID == "txtWhse" && _SysCFLEvent.SelectedObjects != null)
        //                {
        //                    _MasterDBDataSource.SetValue("U_Whse", 0, _SysCFLEvent.SelectedObjects.GetValue("WhsCode", 0).ToString().Trim());
        //                }
        //                if (pVal.ItemUID == "27" && pVal.ColUID == "V_0" && _SysCFLEvent.SelectedObjects != null)
        //                {

        //                    try
        //                    {
        //                        _ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific;
        //                        if (_ComboBoxC.Selected.Value == "Item")
        //                        {
        //                            _Matrix.FlushToDataSource();
        //                            _ChildDBDataSource.SetValue("U_ItemCode", _ChildDBDataSource.Offset, _SysCFLEvent.SelectedObjects.GetValue("ItemCode", 0).ToString().Trim());
        //                            _ChildDBDataSource.SetValue("U_ItemName", _ChildDBDataSource.Offset, _SysCFLEvent.SelectedObjects.GetValue("ItemName", 0).ToString().Trim());
        //                            _ChildDBDataSource.SetValue("U_UOM", _ChildDBDataSource.Offset, _SysCFLEvent.SelectedObjects.GetValue("InvntryUom", 0).ToString().Trim());
        //                            _ChildDBDataSource.SetValue("U_Quantity", _ChildDBDataSource.Offset, "1");
        //                            _ComboBoxB = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_9").Cells.Item(pVal.Row).Specific;
        //                            _Matrix.LoadFromDataSource();
        //                            if (_SysCFLEvent.SelectedObjects.GetValue("IssueMthd", 0).ToString().Trim() == "B")
        //                            {
        //                                _ComboBoxB.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //                            }
        //                            else
        //                            {
        //                                _ComboBoxB.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
        //                            }
        //                            _Matrix.FlushToDataSource();
        //                            _Matrix.LoadFromDataSource();
        //                            AddNewRow();
        //                            _ComboBoxCRow = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row + 1).Specific;
        //                            _ComboBoxCRow.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //                        }
        //                    }
        //                    catch
        //                    {
        //                        Utilities.Application.Company.GetLastError(out ErrCode, out ErrMsg);
        //                    }

        //                }
        //                if (pVal.ItemUID == "27" && pVal.ColUID == "V_3" && _SysCFLEvent.SelectedObjects != null)
        //                {
        //                    try
        //                    {
        //                        _Matrix.FlushToDataSource();
        //                        _ChildDBDataSource.Offset = pVal.Row - 1; // Set the correct row
        //                        _ChildDBDataSource.SetValue("U_Whse", _ChildDBDataSource.Offset, _SysCFLEvent.SelectedObjects.GetValue("WhsCode", 0).ToString().Trim());
        //                        _Matrix.LoadFromDataSource();
        //                        _Matrix.FlushToDataSource();
        //                        //int i = pVal.Row;
        //                        //dynamic cflArg1 = pVal;
        //                        //string UID1 = cflArg1.ChooseFromListUID;
        //                        //dynamic dt1 = cflArg1.SelectedObjects;
        //                        //_Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("27").Specific;
        //                        //((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(i).Specific).Value = dt1.GetValue("WhsCode", 0).ToString();
        //                    }

        //                    catch (Exception ex)
        //                    {
        //                        // Handle the exception, log it, or display an error message.
        //                    }
        //                    chkStatus = true;
        //                }
        //                break;
        //                #endregion
        //        }            
        //    }
        //    #endregion
        //}
        //#endregion

        //#region MENU EVENT
        //public override void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool oBubbleEvent)
        //{
        //    #region BEFORE ACTION
        //    if (pVal.BeforeAction)
        //    {
        //        Form.Freeze(true);
        //        switch (pVal.MenuUID)
        //        {
        //            case Constants.System_Menus.mnu_Duplicate:
        //                {
        //                    _PRNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value;
        //                    Disable_Form_Item();
        //                }
        //                break;
        //        }
        //        Form.Freeze(false);
        //    }
        //    #endregion

        //    #region AFTER ACTION
        //    else
        //    {
        //        try
        //        {
        //            Form.Freeze(true);
        //            switch (pVal.MenuUID)
        //            {
        //                case Constants.System_Menus.mnu_FIND:
        //                    {
        //                        _Form.Items.Item("txtRvnNo").Enabled = true;
                               
        //                    }
        //                    break;
        //                case Constants.System_Menus.mnu_ADD:
        //                    {
        //                        _Form.Items.Item("txtRvnNo").Enabled = false;
        //                    }
        //                    break;
        //                case Constants.System_Menus.mnu_Duplicate:
        //                    {
        //                        _MasterDBDataSource.SetValue("U_RevNum", 0, (Utilities.getMaxColumnValueNumWithItem("@LGBOMH", "U_RevNum", _PRNum)));
        //                        _Form.Items.Item("txtRvnNo").Enabled = false;
        //                        _ComboBoxH.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //                        AddSequenceNumbersToMatrix(_Matrix);
        //                    }
        //                    break;
        //                case Constants.System_Menus.mnu_FIRST:
        //                    _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("27").Specific;
        //                    AddSequenceNumbersToMatrix(_Matrix);
        //                    //Disable_Form_item();
        //                    break;

        //                case Constants.System_Menus.mnu_LAST:
        //                    _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("27").Specific;
        //                    AddSequenceNumbersToMatrix(_Matrix);
        //                    //Disable_Form_item();
        //                    break;

        //                case Constants.System_Menus.mnu_NEXT:
        //                    _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("27").Specific;
        //                    AddSequenceNumbersToMatrix(_Matrix);
        //                    //Disable_Form_item();

        //                    break;

        //                case Constants.System_Menus.mnu_PREVIOUS:
        //                    _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("27").Specific;
        //                    AddSequenceNumbersToMatrix(_Matrix);
        //                    //Disable_Form_item();
        //                    break;

        //            }
        //            if (pVal.MenuUID == "1292" && pVal.BeforeAction == false)
        //            {
        //                AddNewRow();
        //            }
        //        }

        //        catch (Exception ex)
        //        {
        //            Utilities.ShowErrorMessage(ex.Message);
        //        }
        //        finally { Form.Freeze(false); }
        //    }
        //    #endregion
        //}
        //#endregion

        //#region CHOOSE FROM LIST EVENT

        //public override void ChooseFromList_Event()
        //{
        //    try
        //    {
        //        _Form.Freeze(true);
        //        if (this._CFLEvent.ItemUID == "27" && _CFLEvent.ColUID == "V_0")
        //        {
        //            for (int i = 0; i < _CFLEvent.oCFL_SelectedResult.Rows.Count + _CFLEvent.Row - 1; i++)
        //            {
        //                _ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(_Matrix.VisualRowCount).Specific;
        //                _ComboBoxB = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_9").Cells.Item(_Matrix.VisualRowCount).Specific;
        //                if(_ComboBoxC.Selected.Value == "Resource")
        //                {
        //                     int a = i;
        //                    _Matrix.FlushToDataSource();
        //                    _ChildDBDataSource.SetValue("U_ItemCode", _ChildDBDataSource.Offset, _CFLEvent.oCFL_SelectedResult.Rows[i][0].ToString());
        //                    _ChildDBDataSource.SetValue("U_ItemName", _ChildDBDataSource.Offset, _CFLEvent.oCFL_SelectedResult.Rows[i][1].ToString());
        //                    _ChildDBDataSource.SetValue("U_UOM", _ChildDBDataSource.Offset, _CFLEvent.oCFL_SelectedResult.Rows[i][2].ToString());
        //                    _ChildDBDataSource.SetValue("U_Quantity", _ChildDBDataSource.Offset, "1");
        //                    _ComboBoxB.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //                    _Matrix.LoadFromDataSource();
        //                    AddNewRow();
        //                     a = i;
        //                    _ComboBoxCRow = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(_Matrix.VisualRowCount).Specific;
        //                    _ComboBoxCRow.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //                }
        //                else
        //                {
        //                     int a = i;
        //                    _Matrix.FlushToDataSource();
        //                    _ChildDBDataSource.SetValue("U_ItemCode", _ChildDBDataSource.Offset, _CFLEvent.oCFL_SelectedResult.Rows[i][0].ToString());
        //                    _ChildDBDataSource.SetValue("U_ItemName", _ChildDBDataSource.Offset, _CFLEvent.oCFL_SelectedResult.Rows[i][1].ToString());
        //                    _Matrix.LoadFromDataSource();
        //                    AddNewRow();
        //                    a = i;
        //                    _ComboBoxCRow = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(_Matrix.VisualRowCount).Specific;
        //                    _ComboBoxCRow.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        ex.ToString();
        //    }
        //    finally
        //    {
        //        _Form.Freeze(false);
        //    }
        //}

        //#endregion

        //#region  Conditional Reload data in Form

        //private void ConditionalReloadData()
        //{
        //    try
        //    {
        //        _Form.Freeze(true);
        //        _DBData.Query(null);
        //        _Matrix.LoadFromDataSource();
        //        _DBData.Clear();
        //        _Matrix.AddRow(1, _Matrix.RowCount);

        //    }
        //    catch (Exception ex)
        //    {
        //        Utilities.ShowErrorMessage(ex.Message);
        //    }
        //    finally { _Form.Freeze(false); }
        //}

        //#endregion

        //#region FORM DEFAULT
        //internal void FormDefault()
        //{
        //    try
        //    {
        //       _Form.Freeze(true);
        //       _Form.DataBrowser.BrowseBy = "txtEntry";
        //        _Form.Settings.Enabled = true;
        //        UserObjectsMD oUserObjectMD = (UserObjectsMD)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
        //       _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("27").Specific;
        //       _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@LGBOMH");
        //       _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@LGBOMC");
        //       ((SAPbouiCOM.EditText)_Form.Items.Item("txtRvnNo").Specific).Value = (Utilities.getMaxColumnValueNum("@LGBOMH", "U_RevNum"));
        //       ((SAPbouiCOM.EditText)_Form.Items.Item("txtCretBy").Specific).Value = Utilities.Application.SBO_Application.Company.UserName;
        //       ((SAPbouiCOM.EditText)_Form.Items.Item("txtSDate").Specific).Value =  DateTime.Now.ToString("yyyyMMdd");
        //       ((SAPbouiCOM.EditText)_Form.Items.Item("txtVDate").Specific).Value = DateTime.Now.AddDays(30).ToString("yyyyMMdd");
        //       _ComboBoxM = ((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbPList").Specific);
        //       Utilities.FillCombo(ref _ComboBoxM, "select distinct \"ListName\",\"ListNum\" from OPLN order  by \"ListName\" ASC  ");
        //       _ComboBoxH= ((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbStatus").Specific);
        //       _ComboBoxH.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //        AddSequenceNumbersToMatrix(_Matrix);
        //        EnableFromMenu();
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
        //    _Form.Freeze(false);
        //    _Form.PaneLevel = 1;
        //}
        //#endregion

        //#region Add Sequence No
        //private void AddSequenceNumbersToMatrix(Matrix oMatrix)
        //{
        //    //oMatrix.Columns.Item("V_-1").Editable = true;
        //    oMatrix.Columns.Item("V_-1").Visible = true;
        //    for (int i = 1; i <= oMatrix.RowCount; i++)
        //    {
        //        oMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.Value = i.ToString();
        //    }
        //}
        //#endregion

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
        //#endregion(

        //#region FILTER Item Type
        //internal void FilterItemType(string CFLId)
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
        //        _Con.Relationship = BoConditionRelationship.cr_AND;
        //        _Con = _Cons.Add();
        //        _Con.BracketOpenNum = 1;
        //        _Con.Alias = "PrcrmntMtd";
        //        _Con.Operation = BoConditionOperation.co_EQUAL;
        //        _Con.CondVal = "M";
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

        //#region LOAD COMPANY SETTINGS
        //public static DataSet LoadCompanySetting()
        //{
        //    XmlDataDocument xmlDoc = new XmlDataDocument();
        //    DataSet dataSet = new DataSet();
        //    dataSet.ReadXml("AppConfig.xml");
        //    //dataSet.ReadXml("App.Config");
        //    return dataSet;

        //}
        //#endregion

        //#region VALIDATION FUNCTION

        //private bool Validation()
        //{
        //    try
        //    {

               
        //        bool skipRow = false;
        //        for (int i = 1; i <= _Matrix.VisualRowCount-1; i++)
        //        {
                   
        //            if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value.Trim()  == "")
        //            {
        //                Utilities.Message("Type can not left blank", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                return false;
        //            }
        //            if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value.Trim() == "")
        //            {
        //                Utilities.Message("Child Item Code can not left blank", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                return false;
        //            }
        //            if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_7").Cells.Item(i).Specific).Value.Trim() == "" )
        //            {
        //                Utilities.Message("Child Item Name can not left blank", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                return false;
        //            }
        //            if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_6").Cells.Item(i).Specific).Value.Trim()== "" &&((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value.Trim() !="Route Stage")
        //            {
        //                Utilities.Message("Child Quantity can not left blank", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                return false;
        //            }
                   
        //            if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(i).Specific).Value.Trim() == "" && ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value.Trim() != "Route Stage")
        //            {
        //                Utilities.Message("Issue Warehouse can not left blank", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                return false;
        //            }
        //            if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_9").Cells.Item(i).Specific).Value.Trim() == "" && ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value.Trim()!= "Route Stage")
        //            {
        //                Utilities.Message("Issue Method can not left blank", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                return false;
        //            }
        //            Utilities.ExecuteSQL(ref _RSMgItm, "Select T0.\"ManSerNum\", T0.\"ManBtchNum\"  " +
        //                  "From \"OITM\" T0   Where T0.\"ItemCode\" = '" + ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value.Trim() + "' ");
        //            if (_RSMgItm.Fields.Item(0).Value== "Y" || _RSMgItm.Fields.Item(1).Value == "Y")
        //            {
        //                if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_9").Cells.Item(i).Specific).Value.Trim() == "BackFlush")
        //                {
        //                    string itemCode = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value.Trim();
        //                    Utilities.Message("Issue Method can not be BackFlush for Serial/Batch method , Item Code: " + itemCode, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                    return false;
        //                }
        //            }
        //        }
        //        return true;
        //    }
        //    catch
        //    {
        //        throw;
        //    }
        //}

        //#endregion

        //#region Add New Row
        //private void AddNewRow()
        //{
        //    SAPbouiCOM.DBDataSource dbData = default(SAPbouiCOM.DBDataSource);
        //    dbData = (SAPbouiCOM.DBDataSource)_Form.DataSources.DBDataSources.Item(1);
        //    int selRow = _Matrix.GetNextSelectedRow();
        //    dbData.Clear();
        //    _Matrix.AddRow(1, selRow - 1);
        //    AddSequenceNumbersToMatrix(_Matrix);
        //}
        //#endregion

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
        //            _RevNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtRvnNo").Specific).Value;
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
        //                    _BOM.UserFields.Fields.Item("U_RevNum").Value = _RevNum;
        //                    _BOM.UserFields.Fields.Item("U_RevRmrks").Value = _RevRmrks;
        //                    _BOM.PriceList = _PriceList;
                          
        //                    for (int i = 1; i <= _Matrix.RowCount; i++)
        //                    {
        //                        if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value != "")
        //                        {
        //                            if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Route Stage")
        //                            {
        //                                 StageID = (StageID + 1);
        //                                _BOM.Stages.StageEntry = StageID;
        //                                _BOM.Stages.Name = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_7").Cells.Item(i).Specific).Value;
        //                                _BOM.Stages.SequenceNumber = StageID;
        //                                _BOM.Stages.Add();
        //                            }
        //                           else if(((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Resource" || ((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Item")
        //                           {
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
        //                           }
                                   
                                    
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
        //                    _BOM.UserFields.Fields.Item("U_RevNum").Value = _RevNum;
        //                    _BOM.UserFields.Fields.Item("U_RevRmrks").Value = _RevRmrks;
        //                    _BOM.PriceList = _PriceList;
        //                    for (int i = 1; i <= _Matrix.RowCount; i++)
        //                    {
        //                        if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_0").Cells.Item(i).Specific).Value != "")
        //                        {
        //                            if (((SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(i).Specific).Value == "Route Stage")
        //                            {
        //                                 StageID = (StageID + 1);
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

        //#region Disable Form item on Load
        //private void Disable_Form_Item()
        //{
        //    Utilities.ExecuteSQL(ref _RSProd, "Select T0.\"ItemCode\" " +
        //                   "From \"OWOR\" T0   Where T0.\"U_RevNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtRvnNo").Specific).Value + "' " +
        //                   "AND T0.\"ItemCode\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtPrNo").Specific).Value + "'");

        //    Utilities.ExecuteSQL(ref _RSUsr, "Select T0.\"U_StatusAuth\" " +
        //                   "From \"OUSR\" T0   Where T0.\"U_NAME\" = '" + Utilities.Application.SBO_Application.Company.UserName + "' ");
        //    if (_RSProd.RecordCount > 0)
        //    {
        //        _Form.Items.Item("txtQty").Enabled = false;
        //        _Form.Items.Item("txtWhse").Enabled = false;
        //        _Form.Items.Item("txtRvnNo").Enabled = false;
        //        _Form.Items.Item("txtSDate").Enabled = false;
        //        _Form.Items.Item("txtVDate").Enabled = false;
        //        if (_RSUsr.Fields.Item(0).Value!="YES")
        //        {
        //            _Form.Items.Item("cmbStatus").Enabled = false;
        //        }
        //        _Form.Items.Item("txtVRmrks").Enabled = false;
        //        _Form.Items.Item("cmbPList").Enabled = false;
        //        _Form.Items.Item("27").Enabled = false;
        //    }
        //    else
        //    {
        //        _Form.Items.Item("txtQty").Enabled = true;
        //        _Form.Items.Item("txtWhse").Enabled = true;
        //        _Form.Items.Item("txtRvnNo").Enabled = true;
        //        _Form.Items.Item("txtSDate").Enabled = true;
        //        _Form.Items.Item("txtVDate").Enabled = true;
        //        _Form.Items.Item("cmbStatus").Enabled = true;
        //        _Form.Items.Item("txtVRmrks").Enabled = true;
        //        _Form.Items.Item("cmbPList").Enabled = true;
        //        _Form.Items.Item("27").Enabled = true;
        //    }
        //}
        //#endregion

        //#region Enable Form item on Load
        //private void Enable_Form_item()
        //{

        //    _Form.Items.Item("txtQty").Enabled = true;
        //    _Form.Items.Item("txtWhse").Enabled = true;
        //    _Form.Items.Item("txtRvnNo").Enabled = true;
        //    _Form.Items.Item("txtSDate").Enabled = true;
        //    _Form.Items.Item("txtVDate").Enabled = true;
        //    _Form.Items.Item("cmbStatus").Enabled = true;
        //    _Form.Items.Item("txtVRmrks").Enabled = true;
        //    _Form.Items.Item("cmbPList").Enabled = true;
        //    _Form.Items.Item("27").Enabled = true;
        //}
        //#endregion

        //#region Route line setup
        //private void Route_Line_Set(Matrix _Matrix)
        //{
        //    for( int i = 1; i<= _Matrix.RowCount;i++ )
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
    }

}
