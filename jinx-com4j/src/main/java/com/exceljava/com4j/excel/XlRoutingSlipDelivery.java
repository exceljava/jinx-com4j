package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlRoutingSlipDelivery implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlAllAtOnce(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlOneAfterAnother(1),
  ;

  private final int value;
  XlRoutingSlipDelivery(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
