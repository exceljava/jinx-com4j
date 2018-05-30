package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoPresetLightingSoftness implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoPresetLightingSoftnessMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoLightingDim(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoLightingNormal(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoLightingBright(3),
  ;

  private final int value;
  MsoPresetLightingSoftness(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
