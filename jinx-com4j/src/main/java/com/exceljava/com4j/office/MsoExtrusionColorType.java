package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoExtrusionColorType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoExtrusionColorTypeMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoExtrusionColorAutomatic(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoExtrusionColorCustom(2),
  ;

  private final int value;
  MsoExtrusionColorType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
