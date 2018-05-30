package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoFillType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoFillMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoFillSolid(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoFillPatterned(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoFillGradient(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoFillTextured(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoFillBackground(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoFillPicture(6),
  ;

  private final int value;
  MsoFillType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
