package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlOLEType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlOLEControl(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlOLEEmbed(1),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  xlOLELink(0),
  ;

  private final int value;
  XlOLEType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
