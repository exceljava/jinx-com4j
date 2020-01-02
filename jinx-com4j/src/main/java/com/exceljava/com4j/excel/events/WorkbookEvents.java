package com.exceljava.com4j.excel.events;

import com4j.*;

@IID("{00024412-0000-0000-C000-000000000046}")
public abstract class WorkbookEvents {
  // Methods:
  /**
   */

  @DISPID(1923)
  public void open() {
        throw new UnsupportedOperationException();
  }


  /**
   */

  @DISPID(304)
  public void activate() {
        throw new UnsupportedOperationException();
  }


  /**
   */

  @DISPID(1530)
  public void deactivate() {
        throw new UnsupportedOperationException();
  }


  /**
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(1546)
  public void beforeClose(
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param saveAsUI Mandatory boolean parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(1547)
  public void beforeSave(
    boolean saveAsUI,
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(1549)
  public void beforePrint(
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(1550)
  public void newSheet(
    com4j.Com4jObject sh) {
        throw new UnsupportedOperationException();
  }


  /**
   */

  @DISPID(1552)
  public void addinInstall() {
        throw new UnsupportedOperationException();
  }


  /**
   */

  @DISPID(1553)
  public void addinUninstall() {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wn Mandatory com.exceljava.com4j.excel.Window parameter.
   */

  @DISPID(1554)
  public void windowResize(
    com.exceljava.com4j.excel.Window wn) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wn Mandatory com.exceljava.com4j.excel.Window parameter.
   */

  @DISPID(1556)
  public void windowActivate(
    com.exceljava.com4j.excel.Window wn) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param wn Mandatory com.exceljava.com4j.excel.Window parameter.
   */

  @DISPID(1557)
  public void windowDeactivate(
    com.exceljava.com4j.excel.Window wn) {
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
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @DISPID(2158)
  public void pivotTableCloseConnection(
    com.exceljava.com4j.excel.PivotTable target) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @DISPID(2159)
  public void pivotTableOpenConnection(
    com.exceljava.com4j.excel.PivotTable target) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param syncEventType Mandatory com.exceljava.com4j.office.MsoSyncEventType parameter.
   */

  @DISPID(2266)
  public void sync(
    com.exceljava.com4j.office.MsoSyncEventType syncEventType) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param map Mandatory com.exceljava.com4j.excel.XmlMap parameter.
   * @param url Mandatory java.lang.String parameter.
   * @param isRefresh Mandatory boolean parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(2283)
  public void beforeXmlImport(
    com.exceljava.com4j.excel.XmlMap map,
    java.lang.String url,
    boolean isRefresh,
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param map Mandatory com.exceljava.com4j.excel.XmlMap parameter.
   * @param isRefresh Mandatory boolean parameter.
   * @param result Mandatory com.exceljava.com4j.excel.XlXmlImportResult parameter.
   */

  @DISPID(2285)
  public void afterXmlImport(
    com.exceljava.com4j.excel.XmlMap map,
    boolean isRefresh,
    com.exceljava.com4j.excel.XlXmlImportResult result) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param map Mandatory com.exceljava.com4j.excel.XmlMap parameter.
   * @param url Mandatory java.lang.String parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(2287)
  public void beforeXmlExport(
    com.exceljava.com4j.excel.XmlMap map,
    java.lang.String url,
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param map Mandatory com.exceljava.com4j.excel.XmlMap parameter.
   * @param url Mandatory java.lang.String parameter.
   * @param result Mandatory com.exceljava.com4j.excel.XlXmlExportResult parameter.
   */

  @DISPID(2288)
  public void afterXmlExport(
    com.exceljava.com4j.excel.XmlMap map,
    java.lang.String url,
    com.exceljava.com4j.excel.XlXmlExportResult result) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param description Mandatory java.lang.String parameter.
   * @param sheet Mandatory java.lang.String parameter.
   * @param success Mandatory boolean parameter.
   */

  @DISPID(2610)
  public void rowsetComplete(
    java.lang.String description,
    java.lang.String sheet,
    boolean success) {
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
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @DISPID(2899)
  public void sheetPivotTableChangeSync(
    com4j.Com4jObject sh,
    com.exceljava.com4j.excel.PivotTable target) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param success Mandatory boolean parameter.
   */

  @DISPID(2900)
  public void afterSave(
    boolean success) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param ch Mandatory com.exceljava.com4j.excel._Chart parameter.
   */

  @DISPID(2901)
  public void newChart(
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
   * @param changes Mandatory com.exceljava.com4j.excel.ModelChanges parameter.
   */

  @DISPID(3077)
  public void modelChange(
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
   */

  @DISPID(3336)
  public void beforeRemoteChange() {
        throw new UnsupportedOperationException();
  }


  /**
   */

  @DISPID(3337)
  public void afterRemoteChange() {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(3338)
  public void remoteSheetChange(
    com4j.Com4jObject sh,
    com.exceljava.com4j.excel.Range target) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(3339)
  public void remoteNewSheet(
    com4j.Com4jObject sh) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param ch Mandatory com.exceljava.com4j.excel._Chart parameter.
   */

  @DISPID(3340)
  public void remoteNewChart(
    com.exceljava.com4j.excel._Chart ch) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(3341)
  public void remoteSheetBeforeDelete(
    com4j.Com4jObject sh) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @DISPID(3342)
  public void remoteSheetPivotTableUpdate(
    com4j.Com4jObject sh,
    com.exceljava.com4j.excel.PivotTable target) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param sh Mandatory com4j.Com4jObject parameter.
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @DISPID(3343)
  public void remoteSheetPivotTableChangeSync(
    com4j.Com4jObject sh,
    com.exceljava.com4j.excel.PivotTable target) {
        throw new UnsupportedOperationException();
  }


  // Properties:
}
