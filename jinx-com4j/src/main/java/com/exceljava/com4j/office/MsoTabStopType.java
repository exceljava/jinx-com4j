package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoTabStopType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoTabStopMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoTabStopLeft(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoTabStopCenter(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoTabStopRight(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoTabStopDecimal(4),
  ;

  private final int value;
  MsoTabStopType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
