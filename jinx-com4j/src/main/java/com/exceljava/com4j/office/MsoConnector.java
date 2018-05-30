package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoConnector implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoConnectorAnd(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoConnectorOr(2),
  ;

  private final int value;
  MsoConnector(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
