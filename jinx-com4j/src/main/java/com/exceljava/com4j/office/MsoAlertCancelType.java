package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoAlertCancelType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -1
   * </p>
   */
  msoAlertCancelDefault(-1),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoAlertCancelFirst(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoAlertCancelSecond(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoAlertCancelThird(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoAlertCancelFourth(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoAlertCancelFifth(4),
  ;

  private final int value;
  MsoAlertCancelType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
