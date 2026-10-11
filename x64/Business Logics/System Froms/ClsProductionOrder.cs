using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sap.Data.Hana;
using SAPbobsCOM;
using SAPbouiCOM;

namespace Approved_BPCreation
{
    class ClsProductionOrder:Base
    {
        #region VARIABLE DECLARATION

        SAPbouiCOM.Item _ExistItem, _NewItem, _Items;
        private bool flag;
        public SAPbouiCOM.Matrix _Matrix;
        SAPbouiCOM.StaticText _StaticText;
        SAPbouiCOM.Button _Button;
        string _DocEntry = "";
        string _PRNum = string.Empty, _PRName = string.Empty, _RevNum = string.Empty;
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
        private SAPbouiCOM.ComboBox _ComboBoxC;
        private SAPbouiCOM.ComboBox _ComboBoxB;
        private SAPbouiCOM.ComboBox _ComboBoxH;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource = null, _ChildDBDataSource = null, _MasterDBDataSource1 = null;
        private SAPbouiCOM.ChooseFromListCollection _CFLS;
        private SAPbouiCOM.ChooseFromList _CFL;
        private SAPbouiCOM.ChooseFromListCreationParams _CFLC;
        private SAPbouiCOM.Conditions _Cons;
        private SAPbouiCOM.Condition _Con;
        private SAPbouiCOM.Form _UDFForm;
        private int RowID { get; set; }
        private string _Code { get; set; }
        private string _LineId { get; set; }
        private string _ActivityType = string.Empty;
        string _ItemCode = string.Empty, _ItemName = string.Empty;
        private SAPbobsCOM.Recordset _Rs, _RS, _rs, _RSD, _RSBom, _RSPP;
        SAPbouiCOM.ChooseFromListEvent _SysCFLEvent = null;
        SAPbouiCOM.DBDataSource _DBData = null;
        private System.Collections.Generic.List<string> ListOfBlockedID;
        internal string _ServiceID { get; set; }
        internal string _ActivityTypes { get; set; }


        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ClsProductionOrder()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsProductionOrder()
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
                        Form.Freeze(true);
                        Utilities.SerializedMartix(ref _Form, ref _Matrix);
                        Form.Freeze(false);
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
        public override void Item_Event(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            #region Before Action
            if (pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        if (pVal.ItemUID == "txtSlctVer" && ((SAPbouiCOM.EditText)_Form.Items.Item("6").Specific).Value!="")
                        {
                            this.FilterItemType("");
                        }

                        break;
                }
            }

            #endregion

            #region After Action
            //else
            //{
            //    switch (pVal.EventType)
            //    {
            //        #region Item Pressed
            //        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
            //            if (pVal.ItemUID == "BtnProcs")
            //            {
            //                string _Qry = "Select T0.\"U_PRNum\", T0.\"U_PRName\", T0.\"U_Qty\", T0.\"U_Whse\", T0.\"U_UOM\", T0.\"U_RevNum\",T0.\"U_RevRemarks\", " +
            //                              "T1.\"U_Type\",T1.\"U_ItemCode\", T1.\"U_ItemName\", T1.\"U_Quantity\", T1.\"U_Whse\" \"WhseC\" " +
            //                              "From \"@LGBOMH\" T0 inner join \"@LGBOMC\" T1 on T0.\"DocEntry\" = T1.\"DocEntry\"  " +
            //                              "Where T0.\"U_RevNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("U_RevNum").Specific).Value + "' " +
            //                              "AND T0.\"U_PRNum\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("6").Specific).Value + "' " +
            //                              "AND T1.\"U_ItemCode\" != '' ";
            //                Utilities.ExecuteSQL(ref _RS, _Qry);

            //                if (_RS.RecordCount > 0)
            //                {
            //                    // _Form.Freeze(true); // Disable UI updates
            //                    try
            //                    {
            //                        ((SAPbouiCOM.EditText)_Form.Items.Item("78").Specific).Value = _RS.Fields.Item("U_Whse").Value.ToString();
            //                        ((SAPbouiCOM.EditText)_Form.Items.Item("12").Specific).Value = _RS.Fields.Item("U_Qty").Value.ToString();

            //                        for (int i = 1; i <= _RS.RecordCount; i++)
            //                        {
            //                            SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("1880000002").Cells.Item(i).Specific;
            //                            string TypeVal = _RS.Fields.Item("U_Type").Value.ToString().Trim();

            //                            switch (TypeVal)
            //                            {
            //                                case "Route Stage":
            //                                    oComboBox.Select("Route Stage", SAPbouiCOM.BoSearchKey.psk_ByValue);
            //                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("4").Cells.Item(i).Specific).Value = _RS.Fields.Item("U_ItemCode").Value.ToString().Trim();
            //                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("3").Cells.Item(i).Specific).Value = _RS.Fields.Item("U_ItemName").Value.ToString();
            //                                    break;

            //                                case "Resource":
            //                                    oComboBox.Select("Resource", SAPbouiCOM.BoSearchKey.psk_ByValue);
            //                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("4").Cells.Item(i).Specific).Value = _RS.Fields.Item("U_ItemCode").Value.ToString().Trim();
            //                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("3").Cells.Item(i).Specific).Value = _RS.Fields.Item("U_ItemName").Value.ToString();
            //                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("2").Cells.Item(i).Specific).Value = _RS.Fields.Item("U_Quantity").Value.ToString();
            //                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("14").Cells.Item(i).Specific).Value = _RS.Fields.Item("U_Quantity").Value.ToString();
            //                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("10").Cells.Item(i).Specific).Value = _RS.Fields.Item("WhseC").Value.ToString();
            //                                    break;

            //                                default: // Item
            //                                    oComboBox.Select("Item", SAPbouiCOM.BoSearchKey.psk_ByValue);
            //                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("4").Cells.Item(i).Specific).Value = _RS.Fields.Item("U_ItemCode").Value.ToString().Trim();
            //                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("3").Cells.Item(i).Specific).Value = _RS.Fields.Item("U_ItemName").Value.ToString();
            //                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("2").Cells.Item(i).Specific).Value = _RS.Fields.Item("U_Quantity").Value.ToString();
            //                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("14").Cells.Item(i).Specific).Value = _RS.Fields.Item("U_Quantity").Value.ToString();
            //                                    ((SAPbouiCOM.EditText)_Matrix.Columns.Item("10").Cells.Item(i).Specific).Value = _RS.Fields.Item("WhseC").Value.ToString();
            //                                    break;
            //                            }

            //                            _RS.MoveNext();
            //                        }
            //                    }
            //                    finally
            //                    {
            //                        //_Form.Freeze(false); // Re-enable UI updates
            //                    }
            //                }
            //                _Form.Items.Item("BtnProcs").Enabled = false;
            //            }
            //            break;
            //        #endregion

            //        #region Set CFL Values
            //        case BoEventTypes.et_CHOOSE_FROM_LIST:
            //            _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;
            //            if (pVal.ItemUID == "6" && _SysCFLEvent.SelectedObjects != null)
            //            {
            //                string itemCode = _SysCFLEvent.SelectedObjects.GetValue("ItemCode", 0).ToString().Trim();
            //                string orderType = ((SAPbouiCOM.ComboBox)_Form.Items.Item("20").Specific).Value;
            //                _Form.Select();
            //                LoadProductionPC(orderType, itemCode);
            //               // AddSequenceNumbersToMatrix(_Matrix);
            //                _Matrix.AddRow(1, -1);
            //                //AddSequenceNumbersToMatrix(_Matrix);
            //            }
                        
            //            if (pVal.ItemUID == "27" && pVal.ColUID == "V_0" && _SysCFLEvent.SelectedObjects != null)
            //            {

            //                try
            //                {
            //                    _ComboBoxC = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific;
            //                    //if (_ComboBoxC.Selected.Value == "Item")
            //                    //{
            //                    //    _Matrix.FlushToDataSource();
            //                    //    _ChildDBDataSource.SetValue("U_ItemCode", _ChildDBDataSource.Offset, _SysCFLEvent.SelectedObjects.GetValue("ItemCode", 0).ToString().Trim());
            //                    //    _ChildDBDataSource.SetValue("U_ItemName", _ChildDBDataSource.Offset, _SysCFLEvent.SelectedObjects.GetValue("ItemName", 0).ToString().Trim());
            //                    //    _ChildDBDataSource.SetValue("U_UOM", _ChildDBDataSource.Offset, _SysCFLEvent.SelectedObjects.GetValue("InvntryUom", 0).ToString().Trim());
            //                    //    _ChildDBDataSource.SetValue("U_Quantity", _ChildDBDataSource.Offset, "1");
            //                    //    _ComboBoxB = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_9").Cells.Item(pVal.Row).Specific;
            //                    //    _Matrix.LoadFromDataSource();
            //                    //    if (_SysCFLEvent.SelectedObjects.GetValue("IssueMthd", 0).ToString().Trim() == "B")
            //                    //    {
            //                    //        _ComboBoxB.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            //                    //    }
            //                    //    else
            //                    //    {
            //                    //        _ComboBoxB.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
            //                    //    }
            //                    //    _Matrix.FlushToDataSource();
            //                    //    _Matrix.LoadFromDataSource();
            //                    //    AddNewRow();
            //                    //    _ComboBoxCRow = (SAPbouiCOM.ComboBox)_Matrix.Columns.Item("V_8").Cells.Item(pVal.Row + 1).Specific;
            //                    //    _ComboBoxCRow.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            //                    //}
            //                }
            //                catch
            //                {
            //                    Utilities.Application.Company.GetLastError(out ErrCode, out ErrMsg);
            //                }

            //            }
            //            //if (pVal.ItemUID == "27" && pVal.ColUID == "V_3" && _SysCFLEvent.SelectedObjects != null)
            //            //{
            //            //    try
            //            //    {
            //            //        _Matrix.FlushToDataSource();
            //            //        _ChildDBDataSource.Offset = pVal.Row - 1; // Set the correct row
            //            //        _ChildDBDataSource.SetValue("U_Whse", _ChildDBDataSource.Offset, _SysCFLEvent.SelectedObjects.GetValue("WhsCode", 0).ToString().Trim());
            //            //        _Matrix.LoadFromDataSource();
            //            //        _Matrix.FlushToDataSource();
            //            //        //int i = pVal.Row;
            //            //        //dynamic cflArg1 = pVal;
            //            //        //string UID1 = cflArg1.ChooseFromListUID;
            //            //        //dynamic dt1 = cflArg1.SelectedObjects;
            //            //        //_Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("27").Specific;
            //            //        //((SAPbouiCOM.EditText)_Matrix.Columns.Item("V_3").Cells.Item(i).Specific).Value = dt1.GetValue("WhsCode", 0).ToString();
            //            //    }

            //            //    catch (Exception ex)
            //            //    {
            //            //        // Handle the exception, log it, or display an error message.
            //            //    }
            //            //    //chkStatus = true;
            //            //}
            //            break;
            //        #endregion
            //        #endregion

            //        #region Comboselect 
            //        case BoEventTypes.et_COMBO_SELECT:
            //            if (((SAPbouiCOM.ComboBox)_Form.Items.Item("20").Specific).Value == "P")
            //            {
            //                _Form.Items.Item("SlctVer").Visible = true;
            //                _Form.Items.Item("txtSlctVer").Visible = true;
            //                _Form.Items.Item("BtnProcs").Visible = true;
            //            }
            //            break;
            //            #endregion


            //    }
            //}
            #endregion
        }
        
        #region MENU EVENT
        //public override void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool oBubbleEvent)
        //{
        //    #region BEFORE ACTION
        //    if (pVal.BeforeAction)
        //    {

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
                         
        //                    break;

        //                case Constants.System_Menus.mnu_FIRST:

        //                    break;

        //                case Constants.System_Menus.mnu_LAST:

        //                    break;

        //                case Constants.System_Menus.mnu_NEXT:

        //                    break;

        //                case Constants.System_Menus.mnu_PREVIOUS:

        //                    break;
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
        #endregion    

        #region  Conditional Reload data in Form

        private void ConditionalReloadData()
        {
            try
            {
                _Form.Freeze(true);
                _DBData.Query(null);
                _Matrix.LoadFromDataSource();
                _DBData.Clear();
                _Matrix.AddRow(1, _Matrix.RowCount);

            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally { _Form.Freeze(false); }
        }

        #endregion

        #region FORM DEFAULT
        //internal void FormDefault()
        //{
        //    try
        //    {
        //        _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("37").Specific;
        //        _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("OWOR");
        //        _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("WOR1");
        //        _ExistItem = _Form.Items.Item("10");
        //        _NewItem = _Form.Items.Add("SlctVer", SAPbouiCOM.BoFormItemTypes.it_STATIC);
        //        _NewItem.Left = _ExistItem.Left + _ExistItem.Width+2;
        //        _NewItem.Width = _ExistItem.Width -10;
        //        _NewItem.Top = _ExistItem.Top ;
        //        _StaticText = (SAPbouiCOM.StaticText)_NewItem.Specific;
        //        _StaticText.Caption = "Select Version";

        //        _ExistItem = _Form.Items.Item("SlctVer");
        //        _NewItem = _Form.Items.Add("txtSlctVer", SAPbouiCOM.BoFormItemTypes.it_EDIT);
        //        _NewItem.Left = _ExistItem.Left + _ExistItem.Width + 10;
        //        _NewItem.Width = _ExistItem.Width;
        //        _NewItem.Top = _ExistItem.Top;
        //        _EditText = (SAPbouiCOM.EditText)_NewItem.Specific;
        //        SAPbouiCOM.DBDataSource oDataSource = _Form.DataSources.DBDataSources.Add("OWOR");
        //        _EditText.DataBind.SetBound(true, "OWOR", "U_RevNum");

        //        _CFLS = _Form.ChooseFromLists;
        //        _CFLC = Utilities.Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
        //        _CFLC.MultiSelection = false;
        //        _CFLC.ObjectType = "LGBOMH";
        //        _CFLC.UniqueID = "CFL_PRD";
        //        _CFL = _CFLS.Add(_CFLC);
        //        _EditText = ((SAPbouiCOM.EditText)_Form.Items.Item("txtSlctVer").Specific);
        //        //// Assign the ChooseFromList properties to the Edit Text field
        //        _EditText.ChooseFromListUID = "CFL_PRD";
        //        _EditText.ChooseFromListAlias = "U_PRNum";

        //        _ExistItem = _Form.Items.Item("2");
        //        _NewItem = _Form.Items.Add("BtnProcs", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
        //        _NewItem.Left = _ExistItem.Left + _ExistItem.Width + 5;
        //        _NewItem.Width = _ExistItem.Width + 10;
        //        _NewItem.Height = _ExistItem.Height;
        //        _NewItem.Top = _ExistItem.Top;
        //        _Button = (SAPbouiCOM.Button)_NewItem.Specific;
        //        _Button.Caption = "Process";
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
        //    _Form.Items.Item("1000001").Visible = false;
        //    _Form.Items.Item("U_RevNum").Visible = false;
        //    _Form.Items.Item("SlctVer").Visible =false ;
        //    _Form.Items.Item("txtSlctVer").Visible = false;
        //    _Form.Items.Item("BtnProcs").Visible = false;
        //    _Form.Freeze(false);
        //    _Form.PaneLevel = 1;
        //}
        #endregion

        #region FILTER Item Type
        internal void FilterItemType(string CFLId)
        {

            _Form.Freeze(true);
            try
            {
                _Cons = new Conditions();
                _Con = _Cons.Add();
                _Con.BracketOpenNum = 1;
                _Con.Alias = "U_PRNum";
                _Con.Operation = BoConditionOperation.co_EQUAL;
                _Con.CondVal = ((SAPbouiCOM.EditText)_Form.Items.Item("6").Specific).Value;
                _Con.BracketCloseNum = 1;
                _Con.Relationship = BoConditionRelationship.cr_AND;
                _Con = _Cons.Add();
                _Con.BracketOpenNum = 1;
                _Con.Alias = "U_Status";
                _Con.Operation = BoConditionOperation.co_EQUAL;
                _Con.CondVal = "Active";
                _Con.BracketCloseNum = 1;
                _Form.ChooseFromLists.Item("CFL_PRD").SetConditions(_Cons);

            }


            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }

        }
        #endregion

        #region Add New Row
        private void AddNewRow()
        {
            SAPbouiCOM.DBDataSource dbData = default(SAPbouiCOM.DBDataSource);
            dbData = (SAPbouiCOM.DBDataSource)_UDFForm.DataSources.DBDataSources.Item(1);
            int selRow = _Matrix.GetNextSelectedRow();
            dbData.Clear();
            _Matrix.AddRow(1, selRow - 1);
        }
        #endregion

        #region Clear MatrixRow
        private void ClearMatrixRow()
        {
            int rowCount = _Matrix.RowCount;
            for (int i = rowCount; i > 0; i--)
            {
                _Matrix.DeleteRow(i);
            }
            //_Matrix.FlushToDataSource();
            _Matrix.AddRow();

        }
        #endregion

        #region Load Matrix Based on Type
        //private void LoadProductionPC(string orderType, string itemCode)
        //{
        //    try
        //    {
        //        Utilities.Application.SBO_Application.ActivateMenuItem("ProductionOrder");
        //        System.Threading.Thread.Sleep(300);
        //    }
        //    catch
        //    {
        //        // Form might already be open; ignore errors
        //    }

        //    var form = Utilities.Application.SBO_Application.Forms.ActiveForm;
        //    var matrix = (SAPbouiCOM.Matrix)form.Items.Item("37").Specific;

        //    // Clear existing lines
        //    matrix.Clear();

        //    if (orderType == "P")
        //    {
        //        string query = $@"
        //    SELECT
        //        T0.""U_Type"", T0.""U_PRNum"", T0.""U_PRName"", T0.""U_PQty"",
        //        T1.""U_Type"" AS ""LineType"", T1.""U_ItemCode"", T1.""U_ItemName"",
        //        T1.""U_Quantity"", T1.""U_PlannedQty"", T1.""U_Whse"" AS ""LineWhse""
        //    FROM ""@PRODUCTIONPH"" T0
        //    INNER JOIN ""@PRODUCTIONPC"" T1
        //        ON T0.""DocEntry"" = T1.""DocEntry""
        //    WHERE T0.""U_PRNum"" = '{itemCode.Replace("'", "''")}'
        //";

        //        Utilities.ExecuteSQL(ref _RSPP, query);

        //        if (_RSPP.RecordCount > 0)
        //        {
        //            // After Clear(), matrix has one blank row. We need RecordCount rows total.
        //            int needed = _RSPP.RecordCount;
        //            for (int i = 2; i <= needed; i++)
        //            {
        //                matrix.AddRow();
        //            }

        //            _RSPP.MoveFirst();

        //            // Bind header fields
        //            ((SAPbouiCOM.EditText)form.Items.Item("6").Specific).Value = _RSPP.Fields.Item("U_PRNum").Value.ToString().Trim();
        //            ((SAPbouiCOM.EditText)form.Items.Item("12").Specific).Value = _RSPP.Fields.Item("U_PQty").Value.ToString().Trim();

        //            for (int i = 1; i <= needed; i++)
        //            {
        //                string lineType = _RSPP.Fields.Item("LineType").Value.ToString().Trim();
        //                var combo = (SAPbouiCOM.ComboBox)matrix.Columns.Item("1880000002").Cells.Item(i).Specific;
        //                combo.Select(lineType, SAPbouiCOM.BoSearchKey.psk_ByValue);

        //                var txtCode = (SAPbouiCOM.EditText)matrix.Columns.Item("4").Cells.Item(i).Specific;
        //                var txtName = (SAPbouiCOM.EditText)matrix.Columns.Item("3").Cells.Item(i).Specific;

        //                txtCode.Value = _RSPP.Fields.Item("U_ItemCode").Value.ToString().Trim();
        //                txtName.Value = _RSPP.Fields.Item("U_ItemName").Value.ToString().Trim();

        //                if (lineType == "Resource")
        //                {
        //                    ((SAPbouiCOM.EditText)matrix.Columns.Item("2").Cells.Item(i).Specific).Value = _RSPP.Fields.Item("U_Quantity").Value.ToString();
        //                    ((SAPbouiCOM.EditText)matrix.Columns.Item("14").Cells.Item(i).Specific).Value = _RSPP.Fields.Item("U_Quantity").Value.ToString();
        //                    ((SAPbouiCOM.EditText)matrix.Columns.Item("10").Cells.Item(i).Specific).Value = _RSPP.Fields.Item("LineWhse").Value.ToString();
        //                }
        //                else if (lineType != "Route Stage")
        //                {
        //                    ((SAPbouiCOM.EditText)matrix.Columns.Item("14").Cells.Item(i).Specific).Value = _RSPP.Fields.Item("U_PlannedQty").Value.ToString();
        //                }

        //                if (!_RSPP.EoF)
        //                    _RSPP.MoveNext();
        //            }

        //            // Delete any extra blank last row if it exists (item code empty)
        //            int lastRow = matrix.RowCount;
        //            var lastCode = ((SAPbouiCOM.EditText)matrix.Columns.Item("4").Cells.Item(lastRow).Specific).Value;
        //            if (string.IsNullOrWhiteSpace(lastCode))
        //            {
        //                matrix.DeleteRow(lastRow);
        //            }
        //        }
        //    }
        //}

        //private void LoadProductionPC(string orderType, string itemCode)
        //{
        //    var oApp = Utilities.Application.SBO_Application;
        //    try
        //    {
        //        oApp.ActivateMenuItem("ProductionOrder");
        //        System.Threading.Thread.Sleep(300);
        //    }
        //    catch { }

        //    var form = oApp.Forms.ActiveForm;
        //    var matrix = (SAPbouiCOM.Matrix)form.Items.Item("37").Specific;
        //    var lineDS = form.DataSources.DBDataSources.Item("WOR1");

        //    // Clear existing lines
        //    while (matrix.RowCount > 0)
        //    {
        //        matrix.DeleteRow(matrix.RowCount);
        //    }

        //    if (orderType == "P")
        //    {
        //        string query = $@"
        //    SELECT
        //        T1.""U_Type"" AS LineType,
        //                T1.""U_ItemCode"",
        //                T1.""U_ItemName"",
        //                T1.""U_Quantity"",
        //                T1.""U_PlannedQty"",
        //                T1.""U_Whse"" AS LineWhse,
        //                T0.""U_PRNum"",
        //                T0.""U_PQty""
        //            FROM ""@PRODUCTIONPH"" T0
        //            INNER JOIN ""@PRODUCTIONPC"" T1
        //                ON T0.""DocEntry"" = T1.""DocEntry""
        //            WHERE T0.""U_PRNum"" = '{itemCode.Replace("'", "''")}'
        //        ";

        //Utilities.ExecuteSQL(ref _RSPP, query);
        //        if (_RSPP.RecordCount == 0) return;

        //        _RSPP.MoveFirst();
        //        int rowIndex = 0;

        //        while (!_RSPP.EoF)
        //        {
        //            matrix.AddRow();
        //            lineDS.Offset = rowIndex;
        //            lineDS.SetValue("ItemCode", rowIndex, _RSPP.Fields.Item("U_ItemCode").Value.ToString().Trim());
        //            lineDS.SetValue("Quantity", rowIndex, _RSPP.Fields.Item("U_Quantity").Value.ToString().Trim());
        //            lineDS.SetValue("PlannedQty", rowIndex, _RSPP.Fields.Item("U_PlannedQty").Value.ToString().Trim());
        //            lineDS.SetValue("Warehouse", rowIndex, _RSPP.Fields.Item("LineWhse").Value.ToString().Trim());

        //            matrix.SetLineData(rowIndex + 1);

        //            try
        //            {
        //                string lineType = _RSPP.Fields.Item("LineType").Value.ToString().Trim();
        //                if (!string.IsNullOrEmpty(lineType))
        //                {
        //                    var cmb = (SAPbouiCOM.ComboBox)matrix.Columns.Item("1880000002").Cells.Item(rowIndex + 1).Specific;
        //                    cmb.Select(lineType, SAPbouiCOM.BoSearchKey.psk_ByValue);
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                oApp.StatusBar.SetText($"ComboBox error at row {rowIndex + 1}: {ex.Message}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //            }

        //            _RSPP.MoveNext();
        //            rowIndex++;
        //        }

        //        // Header values
        //        _RSPP.MoveFirst();
        //        ((SAPbouiCOM.EditText)form.Items.Item("6").Specific).Value = _RSPP.Fields.Item("U_PRNum").Value.ToString().Trim();
        //        ((SAPbouiCOM.EditText)form.Items.Item("12").Specific).Value = _RSPP.Fields.Item("U_PQty").Value.ToString().Trim();
        //    }
        //}









        #endregion
        #endregion
    }
}
