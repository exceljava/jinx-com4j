package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlDisplayDrawingObjects implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4104
   * </p>
   */
  xlDisplayShapes(-4104),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlHide(3),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlPlaceholders(2),
  ;

  private final int value;
  XlDisplayDrawingObjects(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
