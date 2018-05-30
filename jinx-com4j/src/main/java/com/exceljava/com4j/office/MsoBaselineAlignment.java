package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoBaselineAlignment implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoBaselineAlignMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoBaselineAlignBaseline(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoBaselineAlignTop(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoBaselineAlignCenter(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoBaselineAlignFarEast50(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoBaselineAlignAuto(5),
  ;

  private final int value;
  MsoBaselineAlignment(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
