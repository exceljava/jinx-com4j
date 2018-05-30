package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoArrowheadWidth implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoArrowheadWidthMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoArrowheadNarrow(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoArrowheadWidthMedium(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoArrowheadWide(3),
  ;

  private final int value;
  MsoArrowheadWidth(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
