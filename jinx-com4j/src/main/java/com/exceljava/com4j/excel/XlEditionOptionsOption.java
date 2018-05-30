package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlEditionOptionsOption implements ComEnum {
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlAutomaticUpdate(4),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlCancel(1),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlChangeAttributes(6),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlManualUpdate(5),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlOpenSource(3),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlSelect(3),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlSendPublisher(2),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlUpdateSubscriber(2),
  ;

  private final int value;
  XlEditionOptionsOption(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
