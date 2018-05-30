package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlDataLabelPosition implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4108
   * </p>
   */
  xlLabelPositionCenter(-4108),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  xlLabelPositionAbove(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlLabelPositionBelow(1),
  /**
   * <p>
   * The value of this constant is -4131
   * </p>
   */
  xlLabelPositionLeft(-4131),
  /**
   * <p>
   * The value of this constant is -4152
   * </p>
   */
  xlLabelPositionRight(-4152),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlLabelPositionOutsideEnd(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlLabelPositionInsideEnd(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlLabelPositionInsideBase(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlLabelPositionBestFit(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlLabelPositionMixed(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlLabelPositionCustom(7),
  ;

  private final int value;
  XlDataLabelPosition(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
