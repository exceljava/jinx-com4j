package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoLineStyle implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoLineStyleMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoLineSingle(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoLineThinThin(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoLineThinThick(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoLineThickThin(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoLineThickBetweenThin(5),
  ;

  private final int value;
  MsoLineStyle(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
