package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlLinkInfo implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlEditionDate(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlUpdateState(1),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlLinkInfoStatus(3),
  ;

  private final int value;
  XlLinkInfo(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
