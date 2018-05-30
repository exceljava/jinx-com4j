package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoTextStrike implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoStrikeMixed(-2),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoNoStrike(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoSingleStrike(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoDoubleStrike(2),
  ;

  private final int value;
  MsoTextStrike(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
