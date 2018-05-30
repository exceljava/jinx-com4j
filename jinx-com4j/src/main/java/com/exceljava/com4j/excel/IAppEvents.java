package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024413-0001-0000-C000-000000000046}")
public interface IAppEvents extends Com4jObject {
  // Methods:
  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   */

  @VTID(7)
  void newWorkbook(
    com.exceljava.com4j.excel._Workbook wb);


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @VTID(8)
  void sheetSelectionChange(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh,
    com.exceljava.com4j.excel.Range target);


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @VTID(9)
  void sheetBeforeDoubleClick(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh,
    com.exceljava.com4j.excel.Range target,
    Holder<Boolean> cancel);


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @VTID(10)
  void sheetBeforeRightClick(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh,
    com.exceljava.com4j.excel.Range target,
    Holder<Boolean> cancel);


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @VTID(11)
  void sheetActivate(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh);


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @VTID(12)
  void sheetDeactivate(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh);


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @VTID(13)
  void sheetCalculate(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh);


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @VTID(14)
  void sheetChange(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh,
    com.exceljava.com4j.excel.Range target);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   */

  @VTID(15)
  void workbookOpen(
    com.exceljava.com4j.excel._Workbook wb);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   */

  @VTID(16)
  void workbookActivate(
    com.exceljava.com4j.excel._Workbook wb);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   */

  @VTID(17)
  void workbookDeactivate(
    com.exceljava.com4j.excel._Workbook wb);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @VTID(18)
  void workbookBeforeClose(
    com.exceljava.com4j.excel._Workbook wb,
    Holder<Boolean> cancel);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param saveAsUI Mandatory boolean parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @VTID(19)
  void workbookBeforeSave(
    com.exceljava.com4j.excel._Workbook wb,
    boolean saveAsUI,
    Holder<Boolean> cancel);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @VTID(20)
  void workbookBeforePrint(
    com.exceljava.com4j.excel._Workbook wb,
    Holder<Boolean> cancel);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @VTID(21)
  void workbookNewSheet(
    com.exceljava.com4j.excel._Workbook wb,
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   */

  @VTID(22)
  void workbookAddinInstall(
    com.exceljava.com4j.excel._Workbook wb);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   */

  @VTID(23)
  void workbookAddinUninstall(
    com.exceljava.com4j.excel._Workbook wb);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param wn Mandatory com.exceljava.com4j.excel.Window parameter.
   */

  @VTID(24)
  void windowResize(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.Window wn);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param wn Mandatory com.exceljava.com4j.excel.Window parameter.
   */

  @VTID(25)
  void windowActivate(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.Window wn);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param wn Mandatory com.exceljava.com4j.excel.Window parameter.
   */

  @VTID(26)
  void windowDeactivate(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.Window wn);


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.Hyperlink parameter.
   */

  @VTID(27)
  void sheetFollowHyperlink(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh,
    com.exceljava.com4j.excel.Hyperlink target);


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @VTID(28)
  void sheetPivotTableUpdate(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh,
    com.exceljava.com4j.excel.PivotTable target);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @VTID(29)
  void workbookPivotTableCloseConnection(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.PivotTable target);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @VTID(30)
  void workbookPivotTableOpenConnection(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.PivotTable target);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param syncEventType Mandatory com.exceljava.com4j.office.MsoSyncEventType parameter.
   */

  @VTID(31)
  void workbookSync(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.office.MsoSyncEventType syncEventType);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param map Mandatory com.exceljava.com4j.excel.XmlMap parameter.
   * @param url Mandatory java.lang.String parameter.
   * @param isRefresh Mandatory boolean parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @VTID(32)
  void workbookBeforeXmlImport(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.XmlMap map,
    java.lang.String url,
    boolean isRefresh,
    Holder<Boolean> cancel);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param map Mandatory com.exceljava.com4j.excel.XmlMap parameter.
   * @param isRefresh Mandatory boolean parameter.
   * @param result Mandatory com.exceljava.com4j.excel.XlXmlImportResult parameter.
   */

  @VTID(33)
  void workbookAfterXmlImport(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.XmlMap map,
    boolean isRefresh,
    com.exceljava.com4j.excel.XlXmlImportResult result);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param map Mandatory com.exceljava.com4j.excel.XmlMap parameter.
   * @param url Mandatory java.lang.String parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @VTID(34)
  void workbookBeforeXmlExport(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.XmlMap map,
    java.lang.String url,
    Holder<Boolean> cancel);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param map Mandatory com.exceljava.com4j.excel.XmlMap parameter.
   * @param url Mandatory java.lang.String parameter.
   * @param result Mandatory com.exceljava.com4j.excel.XlXmlExportResult parameter.
   */

  @VTID(35)
  void workbookAfterXmlExport(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.XmlMap map,
    java.lang.String url,
    com.exceljava.com4j.excel.XlXmlExportResult result);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param description Mandatory java.lang.String parameter.
   * @param sheet Mandatory java.lang.String parameter.
   * @param success Mandatory boolean parameter.
   */

  @VTID(36)
  void workbookRowsetComplete(
    com.exceljava.com4j.excel._Workbook wb,
    java.lang.String description,
    java.lang.String sheet,
    boolean success);


  /**
   */

  @VTID(37)
  void afterCalculate();


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param targetPivotTable Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   * @param targetRange Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @VTID(38)
  void sheetPivotTableAfterValueChange(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh,
    com.exceljava.com4j.excel.PivotTable targetPivotTable,
    com.exceljava.com4j.excel.Range targetRange);


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param targetPivotTable Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   * @param valueChangeStart Mandatory int parameter.
   * @param valueChangeEnd Mandatory int parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @VTID(39)
  void sheetPivotTableBeforeAllocateChanges(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh,
    com.exceljava.com4j.excel.PivotTable targetPivotTable,
    int valueChangeStart,
    int valueChangeEnd,
    Holder<Boolean> cancel);


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param targetPivotTable Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   * @param valueChangeStart Mandatory int parameter.
   * @param valueChangeEnd Mandatory int parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @VTID(40)
  void sheetPivotTableBeforeCommitChanges(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh,
    com.exceljava.com4j.excel.PivotTable targetPivotTable,
    int valueChangeStart,
    int valueChangeEnd,
    Holder<Boolean> cancel);


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param targetPivotTable Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   * @param valueChangeStart Mandatory int parameter.
   * @param valueChangeEnd Mandatory int parameter.
   */

  @VTID(41)
  void sheetPivotTableBeforeDiscardChanges(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh,
    com.exceljava.com4j.excel.PivotTable targetPivotTable,
    int valueChangeStart,
    int valueChangeEnd);


  /**
   * @param pvw Mandatory com.exceljava.com4j.excel.ProtectedViewWindow parameter.
   */

  @VTID(42)
  void protectedViewWindowOpen(
    com.exceljava.com4j.excel.ProtectedViewWindow pvw);


  /**
   * @param pvw Mandatory com.exceljava.com4j.excel.ProtectedViewWindow parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @VTID(43)
  void protectedViewWindowBeforeEdit(
    com.exceljava.com4j.excel.ProtectedViewWindow pvw,
    Holder<Boolean> cancel);


  /**
   * @param pvw Mandatory com.exceljava.com4j.excel.ProtectedViewWindow parameter.
   * @param reason Mandatory com.exceljava.com4j.excel.XlProtectedViewCloseReason parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @VTID(44)
  void protectedViewWindowBeforeClose(
    com.exceljava.com4j.excel.ProtectedViewWindow pvw,
    com.exceljava.com4j.excel.XlProtectedViewCloseReason reason,
    Holder<Boolean> cancel);


  /**
   * @param pvw Mandatory com.exceljava.com4j.excel.ProtectedViewWindow parameter.
   */

  @VTID(45)
  void protectedViewWindowResize(
    com.exceljava.com4j.excel.ProtectedViewWindow pvw);


  /**
   * @param pvw Mandatory com.exceljava.com4j.excel.ProtectedViewWindow parameter.
   */

  @VTID(46)
  void protectedViewWindowActivate(
    com.exceljava.com4j.excel.ProtectedViewWindow pvw);


  /**
   * @param pvw Mandatory com.exceljava.com4j.excel.ProtectedViewWindow parameter.
   */

  @VTID(47)
  void protectedViewWindowDeactivate(
    com.exceljava.com4j.excel.ProtectedViewWindow pvw);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param success Mandatory boolean parameter.
   */

  @VTID(48)
  void workbookAfterSave(
    com.exceljava.com4j.excel._Workbook wb,
    boolean success);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param ch Mandatory com.exceljava.com4j.excel._Chart parameter.
   */

  @VTID(49)
  void workbookNewChart(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel._Chart ch);


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @VTID(50)
  void sheetLensGalleryRenderComplete(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh);


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.TableObject parameter.
   */

  @VTID(51)
  void sheetTableUpdate(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh,
    com.exceljava.com4j.excel.TableObject target);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param changes Mandatory com.exceljava.com4j.excel.ModelChanges parameter.
   */

  @VTID(52)
  void workbookModelChange(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.ModelChanges changes);


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @VTID(53)
  void sheetBeforeDelete(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   */

  @VTID(54)
  void workbookBeforeRemoteChange(
    com.exceljava.com4j.excel._Workbook wb);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   */

  @VTID(55)
  void workbookAfterRemoteChange(
    com.exceljava.com4j.excel._Workbook wb);


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @VTID(56)
  void remoteSheetChange(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh,
    com.exceljava.com4j.excel.Range target);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @VTID(57)
  void remoteWorkbookNewSheet(
    com.exceljava.com4j.excel._Workbook wb,
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh);


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param ch Mandatory com.exceljava.com4j.excel._Chart parameter.
   */

  @VTID(58)
  void remoteWorkbookNewChart(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel._Chart ch);


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @VTID(59)
  void remoteSheetBeforeDelete(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh);


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @VTID(60)
  void remoteSheetPivotTableUpdate(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject sh,
    com.exceljava.com4j.excel.PivotTable target);


  // Properties:
}
