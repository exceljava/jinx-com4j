package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoTextureType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoTextureTypeMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoTexturePreset(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoTextureUserDefined(2),
  ;

  private final int value;
  MsoTextureType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
