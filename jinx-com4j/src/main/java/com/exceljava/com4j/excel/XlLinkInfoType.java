package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlLinkInfoType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlLinkInfoOLELinks(2),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlLinkInfoPublishers(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlLinkInfoSubscribers(6),
  ;

  private final int value;
  XlLinkInfoType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
