package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoLastModified implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoLastModifiedYesterday(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoLastModifiedToday(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoLastModifiedLastWeek(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoLastModifiedThisWeek(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoLastModifiedLastMonth(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoLastModifiedThisMonth(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoLastModifiedAnyTime(7),
  ;

  private final int value;
  MsoLastModified(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
