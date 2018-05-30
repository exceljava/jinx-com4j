package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002440F-0001-0000-C000-000000000046}")
public interface IChartEvents extends Com4jObject {
  // Methods:
  /**
   */

  @VTID(7)
  void activate();


  /**
   */

  @VTID(8)
  void deactivate();


  /**
   */

  @VTID(9)
  void resize();


  /**
   * @param button Mandatory int parameter.
   * @param shift Mandatory int parameter.
   * @param x Mandatory int parameter.
   * @param y Mandatory int parameter.
   */

  @VTID(10)
  void mouseDown(
    int button,
    int shift,
    int x,
    int y);


  /**
   * @param button Mandatory int parameter.
   * @param shift Mandatory int parameter.
   * @param x Mandatory int parameter.
   * @param y Mandatory int parameter.
   */

  @VTID(11)
  void mouseUp(
    int button,
    int shift,
    int x,
    int y);


  /**
   * @param button Mandatory int parameter.
   * @param shift Mandatory int parameter.
   * @param x Mandatory int parameter.
   * @param y Mandatory int parameter.
   */

  @VTID(12)
  void mouseMove(
    int button,
    int shift,
    int x,
    int y);


  /**
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @VTID(13)
  void beforeRightClick(
    Holder<Boolean> cancel);


  /**
   */

  @VTID(14)
  void dragPlot();


  /**
   */

  @VTID(15)
  void dragOver();


  /**
   * @param elementID Mandatory int parameter.
   * @param arg1 Mandatory int parameter.
   * @param arg2 Mandatory int parameter.
   * @param cancel Mandatory Holder<Boolean> parameter.
   */

  @VTID(16)
  void beforeDoubleClick(
    int elementID,
    int arg1,
    int arg2,
    Holder<Boolean> cancel);


  /**
   * @param elementID Mandatory int parameter.
   * @param arg1 Mandatory int parameter.
   * @param arg2 Mandatory int parameter.
   */

  @VTID(17)
  void select(
    int elementID,
    int arg1,
    int arg2);


  /**
   * @param seriesIndex Mandatory int parameter.
   * @param pointIndex Mandatory int parameter.
   */

  @VTID(18)
  void seriesChange(
    int seriesIndex,
    int pointIndex);


  /**
   */

  @VTID(19)
  void calculate();


  /**
   */

  @VTID(20)
  void remoteResize();


  /**
   * @param seriesIndex Mandatory int parameter.
   * @param pointIndex Mandatory int parameter.
   */

  @VTID(21)
  void remoteSeriesChange(
    int seriesIndex,
    int pointIndex);


  // Properties:
}
