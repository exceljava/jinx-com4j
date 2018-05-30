package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024411-0001-0000-C000-000000000046}")
public interface IDocEvents extends Com4jObject {
  // Methods:
  /**
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @VTID(7)
  void selectionChange(
    com.exceljava.com4j.excel.Range target);


  /**
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @VTID(8)
  void beforeDoubleClick(
    com.exceljava.com4j.excel.Range target,
    Holder<Boolean> cancel);


  /**
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @VTID(9)
  void beforeRightClick(
    com.exceljava.com4j.excel.Range target,
    Holder<Boolean> cancel);


  /**
   */

  @VTID(10)
  void activate();


  /**
   */

  @VTID(11)
  void deactivate();


  /**
   */

  @VTID(12)
  void calculate();


  /**
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @VTID(13)
  void change(
    com.exceljava.com4j.excel.Range target);


  /**
   * @param target Mandatory com.exceljava.com4j.excel.Hyperlink parameter.
   */

  @VTID(14)
  void followHyperlink(
    com.exceljava.com4j.excel.Hyperlink target);


  /**
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @VTID(15)
  void pivotTableUpdate(
    com.exceljava.com4j.excel.PivotTable target);


  /**
   * @param targetPivotTable Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   * @param targetRange Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @VTID(16)
  void pivotTableAfterValueChange(
    com.exceljava.com4j.excel.PivotTable targetPivotTable,
    com.exceljava.com4j.excel.Range targetRange);


  /**
   * @param targetPivotTable Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   * @param valueChangeStart Mandatory int parameter.
   * @param valueChangeEnd Mandatory int parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @VTID(17)
  void pivotTableBeforeAllocateChanges(
    com.exceljava.com4j.excel.PivotTable targetPivotTable,
    int valueChangeStart,
    int valueChangeEnd,
    Holder<Boolean> cancel);


  /**
   * @param targetPivotTable Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   * @param valueChangeStart Mandatory int parameter.
   * @param valueChangeEnd Mandatory int parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @VTID(18)
  void pivotTableBeforeCommitChanges(
    com.exceljava.com4j.excel.PivotTable targetPivotTable,
    int valueChangeStart,
    int valueChangeEnd,
    Holder<Boolean> cancel);


  /**
   * @param targetPivotTable Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   * @param valueChangeStart Mandatory int parameter.
   * @param valueChangeEnd Mandatory int parameter.
   */

  @VTID(19)
  void pivotTableBeforeDiscardChanges(
    com.exceljava.com4j.excel.PivotTable targetPivotTable,
    int valueChangeStart,
    int valueChangeEnd);


  /**
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @VTID(20)
  void pivotTableChangeSync(
    com.exceljava.com4j.excel.PivotTable target);


  /**
   */

  @VTID(21)
  void lensGalleryRenderComplete();


  /**
   * @param target Mandatory com.exceljava.com4j.excel.TableObject parameter.
   */

  @VTID(22)
  void tableUpdate(
    com.exceljava.com4j.excel.TableObject target);


  /**
   */

  @VTID(23)
  void beforeDelete();


  /**
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @VTID(24)
  void remoteChange(
    com.exceljava.com4j.excel.Range target);


  /**
   */

  @VTID(25)
  void remoteBeforeDelete();


  /**
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @VTID(26)
  void remotePivotTableUpdate(
    com.exceljava.com4j.excel.PivotTable target);


  /**
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @VTID(27)
  void remotePivotTableChangeSync(
    com.exceljava.com4j.excel.PivotTable target);


  // Properties:
}
