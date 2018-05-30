package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoFileFindOptions implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoOptionsNew(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoOptionsAdd(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoOptionsWithin(3),
  ;

  private final int value;
  MsoFileFindOptions(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
