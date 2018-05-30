package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoLineFillType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoLineFillMixed(-2),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoLineFillNone(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoLineFillSolid(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoLineFillPatterned(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoLineFillGradient(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoLineFillTextured(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoLineFillBackground(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoLineFillPicture(6),
  ;

  private final int value;
  MsoLineFillType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
