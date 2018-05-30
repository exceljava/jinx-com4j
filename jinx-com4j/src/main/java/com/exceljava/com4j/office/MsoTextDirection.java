package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoTextDirection implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoTextDirectionMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoTextDirectionLeftToRight(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoTextDirectionRightToLeft(2),
  ;

  private final int value;
  MsoTextDirection(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
