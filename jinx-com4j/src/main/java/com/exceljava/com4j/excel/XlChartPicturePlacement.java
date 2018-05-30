package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlChartPicturePlacement implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlSides(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlEnd(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlEndSides(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlFront(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlFrontSides(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlFrontEnd(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlAllFaces(7),
  ;

  private final int value;
  XlChartPicturePlacement(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
