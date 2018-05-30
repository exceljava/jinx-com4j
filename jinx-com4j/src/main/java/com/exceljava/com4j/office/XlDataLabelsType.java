package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlDataLabelsType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4142
   * </p>
   */
  xlDataLabelsShowNone(-4142),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlDataLabelsShowValue(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlDataLabelsShowPercent(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlDataLabelsShowLabel(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlDataLabelsShowLabelAndPercent(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlDataLabelsShowBubbleSizes(6),
  ;

  private final int value;
  XlDataLabelsType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
