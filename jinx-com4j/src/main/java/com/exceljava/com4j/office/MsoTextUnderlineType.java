package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoTextUnderlineType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoUnderlineMixed(-2),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoNoUnderline(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoUnderlineWords(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoUnderlineSingleLine(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoUnderlineDoubleLine(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoUnderlineHeavyLine(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoUnderlineDottedLine(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoUnderlineDottedHeavyLine(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoUnderlineDashLine(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoUnderlineDashHeavyLine(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoUnderlineDashLongLine(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  msoUnderlineDashLongHeavyLine(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  msoUnderlineDotDashLine(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  msoUnderlineDotDashHeavyLine(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  msoUnderlineDotDotDashLine(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  msoUnderlineDotDotDashHeavyLine(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  msoUnderlineWavyLine(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  msoUnderlineWavyHeavyLine(16),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  msoUnderlineWavyDoubleLine(17),
  ;

  private final int value;
  MsoTextUnderlineType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
