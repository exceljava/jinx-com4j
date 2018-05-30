package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPortugueseReform implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlPortuguesePreReform(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlPortuguesePostReform(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlPortugueseBoth(3),
  ;

  private final int value;
  XlPortugueseReform(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
