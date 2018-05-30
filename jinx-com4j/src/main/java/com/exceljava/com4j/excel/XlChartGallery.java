package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlChartGallery implements ComEnum {
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  xlBuiltIn(21),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  xlUserDefined(22),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  xlAnyGallery(23),
  ;

  private final int value;
  XlChartGallery(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
