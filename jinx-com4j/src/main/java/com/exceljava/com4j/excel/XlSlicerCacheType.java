package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSlicerCacheType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlSlicer(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlTimeline(2),
  ;

  private final int value;
  XlSlicerCacheType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
