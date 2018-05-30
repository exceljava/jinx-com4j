package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoGradientColorType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoGradientColorMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoGradientOneColor(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoGradientTwoColors(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoGradientPresetColors(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoGradientMultiColor(4),
  ;

  private final int value;
  MsoGradientColorType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
