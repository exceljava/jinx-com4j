package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlPivotFieldOrientation implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlColumnField(2),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlDataField(4),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  xlHidden(0),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlPageField(3),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlRowField(1),
  ;

  private final int value;
  XlPivotFieldOrientation(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
