package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlLink implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlExcelLinks(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlOLELinks(2),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlPublishers(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlSubscribers(6),
  ;

  private final int value;
  XlLink(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
