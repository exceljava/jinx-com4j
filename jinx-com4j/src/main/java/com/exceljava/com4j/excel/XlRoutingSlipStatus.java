package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlRoutingSlipStatus implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  xlNotYetRouted(0),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlRoutingComplete(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlRoutingInProgress(1),
  ;

  private final int value;
  XlRoutingSlipStatus(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
