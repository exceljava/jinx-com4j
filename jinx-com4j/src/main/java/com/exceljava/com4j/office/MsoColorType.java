package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoColorType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoColorTypeMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoColorTypeRGB(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoColorTypeScheme(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoColorTypeCMYK(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoColorTypeCMS(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoColorTypeInk(5),
  ;

  private final int value;
  MsoColorType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
