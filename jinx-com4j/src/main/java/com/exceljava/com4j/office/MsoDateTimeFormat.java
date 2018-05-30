package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoDateTimeFormat implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoDateTimeFormatMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoDateTimeMdyy(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoDateTimeddddMMMMddyyyy(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoDateTimedMMMMyyyy(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoDateTimeMMMMdyyyy(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoDateTimedMMMyy(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoDateTimeMMMMyy(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoDateTimeMMyy(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoDateTimeMMddyyHmm(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoDateTimeMMddyyhmmAMPM(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  msoDateTimeHmm(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  msoDateTimeHmmss(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  msoDateTimehmmAMPM(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  msoDateTimehmmssAMPM(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  msoDateTimeFigureOut(14),
  ;

  private final int value;
  MsoDateTimeFormat(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
