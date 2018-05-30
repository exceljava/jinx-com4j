package com.exceljava.com4j.excel.events;

import com4j.*;

@IID("{00024413-0000-0000-C000-000000000046}")
public abstract class AppEvents {
  // Methods:
  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   */

  @DISPID(1565)
  public void newWorkbook(
    com.exceljava.com4j.excel._Workbook wb) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(1558)
  public void sheetSelectionChange(
    com4j.Com4jObject sh,
    com.exceljava.com4j.excel.Range target) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(1559)
  public void sheetBeforeDoubleClick(
    com4j.Com4jObject sh,
    com.exceljava.com4j.excel.Range target,
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(1560)
  public void sheetBeforeRightClick(
    com4j.Com4jObject sh,
    com.exceljava.com4j.excel.Range target,
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(1561)
  public void sheetActivate(
    com4j.Com4jObject sh) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(1562)
  public void sheetDeactivate(
    com4j.Com4jObject sh) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(1563)
  public void sheetCalculate(
    com4j.Com4jObject sh) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(1564)
  public void sheetChange(
    com4j.Com4jObject sh,
    com.exceljava.com4j.excel.Range target) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   */

  @DISPID(1567)
  public void workbookOpen(
    com.exceljava.com4j.excel._Workbook wb) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   */

  @DISPID(1568)
  public void workbookActivate(
    com.exceljava.com4j.excel._Workbook wb) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   */

  @DISPID(1569)
  public void workbookDeactivate(
    com.exceljava.com4j.excel._Workbook wb) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(1570)
  public void workbookBeforeClose(
    com.exceljava.com4j.excel._Workbook wb,
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param saveAsUI Mandatory boolean parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(1571)
  public void workbookBeforeSave(
    com.exceljava.com4j.excel._Workbook wb,
    boolean saveAsUI,
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(1572)
  public void workbookBeforePrint(
    com.exceljava.com4j.excel._Workbook wb,
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(1573)
  public void workbookNewSheet(
    com.exceljava.com4j.excel._Workbook wb,
    com4j.Com4jObject sh) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   */

  @DISPID(1574)
  public void workbookAddinInstall(
    com.exceljava.com4j.excel._Workbook wb) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   */

  @DISPID(1575)
  public void workbookAddinUninstall(
    com.exceljava.com4j.excel._Workbook wb) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param wn Mandatory com.exceljava.com4j.excel.Window parameter.
   */

  @DISPID(1554)
  public void windowResize(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.Window wn) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param wn Mandatory com.exceljava.com4j.excel.Window parameter.
   */

  @DISPID(1556)
  public void windowActivate(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.Window wn) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param wn Mandatory com.exceljava.com4j.excel.Window parameter.
   */

  @DISPID(1557)
  public void windowDeactivate(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.Window wn) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.Hyperlink parameter.
   */

  @DISPID(1854)
  public void sheetFollowHyperlink(
    com4j.Com4jObject sh,
    com.exceljava.com4j.excel.Hyperlink target) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @DISPID(2157)
  public void sheetPivotTableUpdate(
    com4j.Com4jObject sh,
    com.exceljava.com4j.excel.PivotTable target) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @DISPID(2160)
  public void workbookPivotTableCloseConnection(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.PivotTable target) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @DISPID(2161)
  public void workbookPivotTableOpenConnection(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.PivotTable target) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param syncEventType Mandatory com.exceljava.com4j.office.MsoSyncEventType parameter.
   */

  @DISPID(2289)
  public void workbookSync(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.office.MsoSyncEventType syncEventType) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param map Mandatory com.exceljava.com4j.excel.XmlMap parameter.
   * @param url Mandatory java.lang.String parameter.
   * @param isRefresh Mandatory boolean parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(2290)
  public void workbookBeforeXmlImport(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.XmlMap map,
    java.lang.String url,
    boolean isRefresh,
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param map Mandatory com.exceljava.com4j.excel.XmlMap parameter.
   * @param isRefresh Mandatory boolean parameter.
   * @param result Mandatory com.exceljava.com4j.excel.XlXmlImportResult parameter.
   */

  @DISPID(2291)
  public void workbookAfterXmlImport(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.XmlMap map,
    boolean isRefresh,
    com.exceljava.com4j.excel.XlXmlImportResult result) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param map Mandatory com.exceljava.com4j.excel.XmlMap parameter.
   * @param url Mandatory java.lang.String parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(2292)
  public void workbookBeforeXmlExport(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.XmlMap map,
    java.lang.String url,
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param map Mandatory com.exceljava.com4j.excel.XmlMap parameter.
   * @param url Mandatory java.lang.String parameter.
   * @param result Mandatory com.exceljava.com4j.excel.XlXmlExportResult parameter.
   */

  @DISPID(2293)
  public void workbookAfterXmlExport(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.XmlMap map,
    java.lang.String url,
    com.exceljava.com4j.excel.XlXmlExportResult result) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param description Mandatory java.lang.String parameter.
   * @param sheet Mandatory java.lang.String parameter.
   * @param success Mandatory boolean parameter.
   */

  @DISPID(2611)
  public void workbookRowsetComplete(
    com.exceljava.com4j.excel._Workbook wb,
    java.lang.String description,
    java.lang.String sheet,
    boolean success) {
        throw new UnsupportedOperationException();
  }


  /**
   */

  @DISPID(2612)
  public void afterCalculate() {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param targetPivotTable Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   * @param targetRange Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(2895)
  public void sheetPivotTableAfterValueChange(
    com4j.Com4jObject sh,
    com.exceljava.com4j.excel.PivotTable targetPivotTable,
    com.exceljava.com4j.excel.Range targetRange) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param targetPivotTable Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   * @param valueChangeStart Mandatory int parameter.
   * @param valueChangeEnd Mandatory int parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(2896)
  public void sheetPivotTableBeforeAllocateChanges(
    com4j.Com4jObject sh,
    com.exceljava.com4j.excel.PivotTable targetPivotTable,
    int valueChangeStart,
    int valueChangeEnd,
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param targetPivotTable Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   * @param valueChangeStart Mandatory int parameter.
   * @param valueChangeEnd Mandatory int parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(2897)
  public void sheetPivotTableBeforeCommitChanges(
    com4j.Com4jObject sh,
    com.exceljava.com4j.excel.PivotTable targetPivotTable,
    int valueChangeStart,
    int valueChangeEnd,
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param targetPivotTable Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   * @param valueChangeStart Mandatory int parameter.
   * @param valueChangeEnd Mandatory int parameter.
   */

  @DISPID(2898)
  public void sheetPivotTableBeforeDiscardChanges(
    com4j.Com4jObject sh,
    com.exceljava.com4j.excel.PivotTable targetPivotTable,
    int valueChangeStart,
    int valueChangeEnd) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param pvw Mandatory com.exceljava.com4j.excel.ProtectedViewWindow parameter.
   */

  @DISPID(2903)
  public void protectedViewWindowOpen(
    com.exceljava.com4j.excel.ProtectedViewWindow pvw) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param pvw Mandatory com.exceljava.com4j.excel.ProtectedViewWindow parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(2905)
  public void protectedViewWindowBeforeEdit(
    com.exceljava.com4j.excel.ProtectedViewWindow pvw,
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param pvw Mandatory com.exceljava.com4j.excel.ProtectedViewWindow parameter.
   * @param reason Mandatory com.exceljava.com4j.excel.XlProtectedViewCloseReason parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(2906)
  public void protectedViewWindowBeforeClose(
    com.exceljava.com4j.excel.ProtectedViewWindow pvw,
    com.exceljava.com4j.excel.XlProtectedViewCloseReason reason,
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param pvw Mandatory com.exceljava.com4j.excel.ProtectedViewWindow parameter.
   */

  @DISPID(2908)
  public void protectedViewWindowResize(
    com.exceljava.com4j.excel.ProtectedViewWindow pvw) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param pvw Mandatory com.exceljava.com4j.excel.ProtectedViewWindow parameter.
   */

  @DISPID(2909)
  public void protectedViewWindowActivate(
    com.exceljava.com4j.excel.ProtectedViewWindow pvw) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param pvw Mandatory com.exceljava.com4j.excel.ProtectedViewWindow parameter.
   */

  @DISPID(2910)
  public void protectedViewWindowDeactivate(
    com.exceljava.com4j.excel.ProtectedViewWindow pvw) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param success Mandatory boolean parameter.
   */

  @DISPID(2911)
  public void workbookAfterSave(
    com.exceljava.com4j.excel._Workbook wb,
    boolean success) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param ch Mandatory com.exceljava.com4j.excel._Chart parameter.
   */

  @DISPID(2912)
  public void workbookNewChart(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel._Chart ch) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(3075)
  public void sheetLensGalleryRenderComplete(
    com4j.Com4jObject sh) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.TableObject parameter.
   */

  @DISPID(3076)
  public void sheetTableUpdate(
    com4j.Com4jObject sh,
    com.exceljava.com4j.excel.TableObject target) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param changes Mandatory com.exceljava.com4j.excel.ModelChanges parameter.
   */

  @DISPID(3080)
  public void workbookModelChange(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel.ModelChanges changes) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(3079)
  public void sheetBeforeDelete(
    com4j.Com4jObject sh) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   */

  @DISPID(3293)
  public void workbookBeforeRemoteChange(
    com.exceljava.com4j.excel._Workbook wb) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   */

  @DISPID(3294)
  public void workbookAfterRemoteChange(
    com.exceljava.com4j.excel._Workbook wb) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(3287)
  public void remoteSheetChange(
    com4j.Com4jObject sh,
    com.exceljava.com4j.excel.Range target) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(3295)
  public void remoteWorkbookNewSheet(
    com.exceljava.com4j.excel._Workbook wb,
    com4j.Com4jObject sh) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wb Mandatory com.exceljava.com4j.excel._Workbook parameter.
   * @param ch Mandatory com.exceljava.com4j.excel._Chart parameter.
   */

  @DISPID(3296)
  public void remoteWorkbookNewChart(
    com.exceljava.com4j.excel._Workbook wb,
    com.exceljava.com4j.excel._Chart ch) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(3290)
  public void remoteSheetBeforeDelete(
    com4j.Com4jObject sh) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @DISPID(3291)
  public void remoteSheetPivotTableUpdate(
    com4j.Com4jObject sh,
    com.exceljava.com4j.excel.PivotTable target) {
        throw new UnsupportedOperationException();
  }


  // Properties:
}
