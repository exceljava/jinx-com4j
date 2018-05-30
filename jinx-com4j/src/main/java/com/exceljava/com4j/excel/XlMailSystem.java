package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlMailSystem implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlMAPI(1),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  xlNoMailSystem(0),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlPowerTalk(2),
  ;

  private final int value;
  XlMailSystem(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
