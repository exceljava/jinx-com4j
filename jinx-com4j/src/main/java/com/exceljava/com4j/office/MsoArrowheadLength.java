package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoArrowheadLength implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoArrowheadLengthMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoArrowheadShort(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoArrowheadLengthMedium(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoArrowheadLong(3),
  ;

  private final int value;
  MsoArrowheadLength(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
