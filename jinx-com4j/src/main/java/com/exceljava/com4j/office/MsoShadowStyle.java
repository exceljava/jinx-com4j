package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoShadowStyle implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoShadowStyleMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoShadowStyleInnerShadow(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoShadowStyleOuterShadow(2),
  ;

  private final int value;
  MsoShadowStyle(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
