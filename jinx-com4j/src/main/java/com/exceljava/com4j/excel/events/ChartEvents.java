package com.exceljava.com4j.excel.events;

import com4j.*;

@IID("{0002440F-0000-0000-C000-000000000046}")
public abstract class ChartEvents {
  // Methods:
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

  @DISPID(256)
  public void resize() {
        throw new UnsupportedOperationException();
  }


  /**
   * @param button Mandatory int parameter.
   * @param shift Mandatory int parameter.
   * @param x Mandatory int parameter.
   * @param y Mandatory int parameter.
   */

  @DISPID(1531)
  public void mouseDown(
    int button,
    int shift,
    int x,
    int y) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param button Mandatory int parameter.
   * @param shift Mandatory int parameter.
   * @param x Mandatory int parameter.
   * @param y Mandatory int parameter.
   */

  @DISPID(1532)
  public void mouseUp(
    int button,
    int shift,
    int x,
    int y) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param button Mandatory int parameter.
   * @param shift Mandatory int parameter.
   * @param x Mandatory int parameter.
   * @param y Mandatory int parameter.
   */

  @DISPID(1533)
  public void mouseMove(
    int button,
    int shift,
    int x,
    int y) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(1534)
  public void beforeRightClick(
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   */

  @DISPID(1535)
  public void dragPlot() {
        throw new UnsupportedOperationException();
  }


  /**
   */

  @DISPID(1536)
  public void dragOver() {
        throw new UnsupportedOperationException();
  }


  /**
   * @param elementID Mandatory int parameter.
   * @param arg1 Mandatory int parameter.
   * @param arg2 Mandatory int parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @DISPID(1537)
  public void beforeDoubleClick(
    int elementID,
    int arg1,
    int arg2,
    Holder<Boolean> cancel) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param elementID Mandatory int parameter.
   * @param arg1 Mandatory int parameter.
   * @param arg2 Mandatory int parameter.
   */

  @DISPID(235)
  public void select(
    int elementID,
    int arg1,
    int arg2) {
        throw new UnsupportedOperationException();
  }


  /**
   * @param seriesIndex Mandatory int parameter.
   * @param pointIndex Mandatory int parameter.
   */

  @DISPID(1538)
  public void seriesChange(
    int seriesIndex,
    int pointIndex) {
        throw new UnsupportedOperationException();
  }


  /**
   */

  @DISPID(279)
  public void calculate() {
        throw new UnsupportedOperationException();
  }


  /**
   */

  @DISPID(3279)
  public void remoteResize() {
        throw new UnsupportedOperationException();
  }


  /**
   * @param seriesIndex Mandatory int parameter.
   * @param pointIndex Mandatory int parameter.
   */

  @DISPID(3280)
  public void remoteSeriesChange(
    int seriesIndex,
    int pointIndex) {
        throw new UnsupportedOperationException();
  }


  // Properties:
}
