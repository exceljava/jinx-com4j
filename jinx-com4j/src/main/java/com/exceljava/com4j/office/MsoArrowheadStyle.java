package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoArrowheadStyle implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoArrowheadStyleMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoArrowheadNone(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoArrowheadTriangle(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoArrowheadOpen(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoArrowheadStealth(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoArrowheadDiamond(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoArrowheadOval(6),
  ;

  private final int value;
  MsoArrowheadStyle(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
