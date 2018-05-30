package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoMixedType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 32768
   * </p>
   */
  msoIntegerMixed(32768),
  /**
   * <p>
   * The value of this constant is -2147483648
   * </p>
   */
  msoSingleMixed(-2147483648),
  ;

  private final int value;
  MsoMixedType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
