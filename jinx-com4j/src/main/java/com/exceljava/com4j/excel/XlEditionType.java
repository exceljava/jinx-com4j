package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlEditionType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlPublisher(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlSubscriber(2),
  ;

  private final int value;
  XlEditionType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
