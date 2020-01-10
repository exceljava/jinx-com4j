package com.exceljava.com4j.excel.events;

import com4j.*;

@IID("{00024411-0000-0000-C000-000000000046}")
public abstract class DocEvents {
  // Methods:
  /**
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(1543)
  public void selectionChange(
    com.exceljava.com4j.excel.Range target) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(1537)
  public void beforeDoubleClick(
    com.exceljava.com4j.excel.Range target,
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(1534)
  public void beforeRightClick(
    com.exceljava.com4j.excel.Range target,
    Holder<Boolean> cancel) {
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
   */

  @DISPID(279)
  public void calculate() {
        throw new UnsupportedOperationException();
  }


  /**
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(1545)
  public void change(
    com.exceljava.com4j.excel.Range target) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param target Mandatory com.exceljava.com4j.excel.Hyperlink parameter.
   */

  @DISPID(1470)
  public void followHyperlink(
    com.exceljava.com4j.excel.Hyperlink target) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @DISPID(2156)
  public void pivotTableUpdate(
    com.exceljava.com4j.excel.PivotTable target) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param targetPivotTable Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   * @param targetRange Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(2886)
  public void pivotTableAfterValueChange(
    com.exceljava.com4j.excel.PivotTable targetPivotTable,
    com.exceljava.com4j.excel.Range targetRange) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param targetPivotTable Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   * @param valueChangeStart Mandatory int parameter.
   * @param valueChangeEnd Mandatory int parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(2889)
  public void pivotTableBeforeAllocateChanges(
    com.exceljava.com4j.excel.PivotTable targetPivotTable,
    int valueChangeStart,
    int valueChangeEnd,
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param targetPivotTable Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   * @param valueChangeStart Mandatory int parameter.
   * @param valueChangeEnd Mandatory int parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(2892)
  public void pivotTableBeforeCommitChanges(
    com.exceljava.com4j.excel.PivotTable targetPivotTable,
    int valueChangeStart,
    int valueChangeEnd,
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param targetPivotTable Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   * @param valueChangeStart Mandatory int parameter.
   * @param valueChangeEnd Mandatory int parameter.
   */

  @DISPID(2893)
  public void pivotTableBeforeDiscardChanges(
    com.exceljava.com4j.excel.PivotTable targetPivotTable,
    int valueChangeStart,
    int valueChangeEnd) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @DISPID(2894)
  public void pivotTableChangeSync(
    com.exceljava.com4j.excel.PivotTable target) {
        throw new UnsupportedOperationException();
  }


  /**
   */

  @DISPID(3072)
  public void lensGalleryRenderComplete() {
        throw new UnsupportedOperationException();
  }


  /**
   * @param target Mandatory com.exceljava.com4j.excel.TableObject parameter.
   */

  @DISPID(3073)
  public void tableUpdate(
    com.exceljava.com4j.excel.TableObject target) {
        throw new UnsupportedOperationException();
  }


  /**
   */

  @DISPID(3074)
  public void beforeDelete() {
        throw new UnsupportedOperationException();
  }


  /**
   * @param target Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(3332)
  public void remoteChange(
    com.exceljava.com4j.excel.Range target) {
        throw new UnsupportedOperationException();
  }


  /**
   */

  @DISPID(3333)
  public void remoteBeforeDelete() {
        throw new UnsupportedOperationException();
  }


  /**
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @DISPID(3334)
  public void remotePivotTableUpdate(
    com.exceljava.com4j.excel.PivotTable target) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param target Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @DISPID(3335)
  public void remotePivotTableChangeSync(
    com.exceljava.com4j.excel.PivotTable target) {
        throw new UnsupportedOperationException();
  }


  // Properties:
}
