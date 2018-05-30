package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoIconType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoIconNone(0),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoIconAlert(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoIconTip(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoIconAlertInfo(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoIconAlertWarning(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoIconAlertQuery(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoIconAlertCritical(7),
  ;

  private final int value;
  MsoIconType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
