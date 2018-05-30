package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoVerticalAnchor implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoVerticalAnchorMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoAnchorTop(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoAnchorTopBaseline(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoAnchorMiddle(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoAnchorBottom(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoAnchorBottomBaseLine(5),
  ;

  private final int value;
  MsoVerticalAnchor(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
