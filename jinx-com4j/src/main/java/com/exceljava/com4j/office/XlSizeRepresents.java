package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlSizeRepresents implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlSizeIsWidth(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlSizeIsArea(1),
  ;

  private final int value;
  XlSizeRepresents(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
