package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoTextureAlignment implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoTextureAlignmentMixed(-2),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoTextureTopLeft(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoTextureTop(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoTextureTopRight(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoTextureLeft(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoTextureCenter(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoTextureRight(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoTextureBottomLeft(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoTextureBottom(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoTextureBottomRight(8),
  ;

  private final int value;
  MsoTextureAlignment(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
