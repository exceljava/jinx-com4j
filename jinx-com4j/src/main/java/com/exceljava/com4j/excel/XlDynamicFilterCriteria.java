package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlDynamicFilterCriteria implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlFilterToday(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlFilterYesterday(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlFilterTomorrow(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlFilterThisWeek(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlFilterLastWeek(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlFilterNextWeek(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlFilterThisMonth(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlFilterLastMonth(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlFilterNextMonth(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlFilterThisQuarter(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlFilterLastQuarter(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  xlFilterNextQuarter(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  xlFilterThisYear(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  xlFilterLastYear(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  xlFilterNextYear(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  xlFilterYearToDate(16),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  xlFilterAllDatesInPeriodQuarter1(17),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  xlFilterAllDatesInPeriodQuarter2(18),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  xlFilterAllDatesInPeriodQuarter3(19),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  xlFilterAllDatesInPeriodQuarter4(20),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  xlFilterAllDatesInPeriodJanuary(21),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  xlFilterAllDatesInPeriodFebruray(22),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  xlFilterAllDatesInPeriodMarch(23),
  /**
   * <p>
   * The value of this constant is 24
   * </p>
   */
  xlFilterAllDatesInPeriodApril(24),
  /**
   * <p>
   * The value of this constant is 25
   * </p>
   */
  xlFilterAllDatesInPeriodMay(25),
  /**
   * <p>
   * The value of this constant is 26
   * </p>
   */
  xlFilterAllDatesInPeriodJune(26),
  /**
   * <p>
   * The value of this constant is 27
   * </p>
   */
  xlFilterAllDatesInPeriodJuly(27),
  /**
   * <p>
   * The value of this constant is 28
   * </p>
   */
  xlFilterAllDatesInPeriodAugust(28),
  /**
   * <p>
   * The value of this constant is 29
   * </p>
   */
  xlFilterAllDatesInPeriodSeptember(29),
  /**
   * <p>
   * The value of this constant is 30
   * </p>
   */
  xlFilterAllDatesInPeriodOctober(30),
  /**
   * <p>
   * The value of this constant is 31
   * </p>
   */
  xlFilterAllDatesInPeriodNovember(31),
  /**
   * <p>
   * The value of this constant is 32
   * </p>
   */
  xlFilterAllDatesInPeriodDecember(32),
  /**
   * <p>
   * The value of this constant is 33
   * </p>
   */
  xlFilterAboveAverage(33),
  /**
   * <p>
   * The value of this constant is 34
   * </p>
   */
  xlFilterBelowAverage(34),
  ;

  private final int value;
  XlDynamicFilterCriteria(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
