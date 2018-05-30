package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlVAlign implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4107
   * </p>
   */
  xlVAlignBottom(-4107),
  /**
   * <p>
   * The value of this constant is -4108
   * </p>
   */
  xlVAlignCenter(-4108),
  /**
   * <p>
   * The value of this constant is -4117
   * </p>
   */
  xlVAlignDistributed(-4117),
  /**
   * <p>
   * The value of this constant is -4130
   * </p>
   */
  xlVAlignJustify(-4130),
  /**
   * <p>
   * The value of this constant is -4160
   * </p>
   */
  xlVAlignTop(-4160),
  ;

  private final int value;
  XlVAlign(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
