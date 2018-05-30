package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPivotFilterType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlTopCount(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlBottomCount(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlTopPercent(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlBottomPercent(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlTopSum(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlBottomSum(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlValueEquals(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlValueDoesNotEqual(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlValueIsGreaterThan(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlValueIsGreaterThanOrEqualTo(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlValueIsLessThan(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  xlValueIsLessThanOrEqualTo(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  xlValueIsBetween(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  xlValueIsNotBetween(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  xlCaptionEquals(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  xlCaptionDoesNotEqual(16),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  xlCaptionBeginsWith(17),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  xlCaptionDoesNotBeginWith(18),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  xlCaptionEndsWith(19),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  xlCaptionDoesNotEndWith(20),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  xlCaptionContains(21),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  xlCaptionDoesNotContain(22),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  xlCaptionIsGreaterThan(23),
  /**
   * <p>
   * The value of this constant is 24
   * </p>
   */
  xlCaptionIsGreaterThanOrEqualTo(24),
  /**
   * <p>
   * The value of this constant is 25
   * </p>
   */
  xlCaptionIsLessThan(25),
  /**
   * <p>
   * The value of this constant is 26
   * </p>
   */
  xlCaptionIsLessThanOrEqualTo(26),
  /**
   * <p>
   * The value of this constant is 27
   * </p>
   */
  xlCaptionIsBetween(27),
  /**
   * <p>
   * The value of this constant is 28
   * </p>
   */
  xlCaptionIsNotBetween(28),
  /**
   * <p>
   * The value of this constant is 29
   * </p>
   */
  xlSpecificDate(29),
  /**
   * <p>
   * The value of this constant is 30
   * </p>
   */
  xlNotSpecificDate(30),
  /**
   * <p>
   * The value of this constant is 31
   * </p>
   */
  xlBefore(31),
  /**
   * <p>
   * The value of this constant is 32
   * </p>
   */
  xlBeforeOrEqualTo(32),
  /**
   * <p>
   * The value of this constant is 33
   * </p>
   */
  xlAfter(33),
  /**
   * <p>
   * The value of this constant is 34
   * </p>
   */
  xlAfterOrEqualTo(34),
  /**
   * <p>
   * The value of this constant is 35
   * </p>
   */
  xlDateBetween(35),
  /**
   * <p>
   * The value of this constant is 36
   * </p>
   */
  xlDateNotBetween(36),
  /**
   * <p>
   * The value of this constant is 37
   * </p>
   */
  xlDateTomorrow(37),
  /**
   * <p>
   * The value of this constant is 38
   * </p>
   */
  xlDateToday(38),
  /**
   * <p>
   * The value of this constant is 39
   * </p>
   */
  xlDateYesterday(39),
  /**
   * <p>
   * The value of this constant is 40
   * </p>
   */
  xlDateNextWeek(40),
  /**
   * <p>
   * The value of this constant is 41
   * </p>
   */
  xlDateThisWeek(41),
  /**
   * <p>
   * The value of this constant is 42
   * </p>
   */
  xlDateLastWeek(42),
  /**
   * <p>
   * The value of this constant is 43
   * </p>
   */
  xlDateNextMonth(43),
  /**
   * <p>
   * The value of this constant is 44
   * </p>
   */
  xlDateThisMonth(44),
  /**
   * <p>
   * The value of this constant is 45
   * </p>
   */
  xlDateLastMonth(45),
  /**
   * <p>
   * The value of this constant is 46
   * </p>
   */
  xlDateNextQuarter(46),
  /**
   * <p>
   * The value of this constant is 47
   * </p>
   */
  xlDateThisQuarter(47),
  /**
   * <p>
   * The value of this constant is 48
   * </p>
   */
  xlDateLastQuarter(48),
  /**
   * <p>
   * The value of this constant is 49
   * </p>
   */
  xlDateNextYear(49),
  /**
   * <p>
   * The value of this constant is 50
   * </p>
   */
  xlDateThisYear(50),
  /**
   * <p>
   * The value of this constant is 51
   * </p>
   */
  xlDateLastYear(51),
  /**
   * <p>
   * The value of this constant is 52
   * </p>
   */
  xlYearToDate(52),
  /**
   * <p>
   * The value of this constant is 53
   * </p>
   */
  xlAllDatesInPeriodQuarter1(53),
  /**
   * <p>
   * The value of this constant is 54
   * </p>
   */
  xlAllDatesInPeriodQuarter2(54),
  /**
   * <p>
   * The value of this constant is 55
   * </p>
   */
  xlAllDatesInPeriodQuarter3(55),
  /**
   * <p>
   * The value of this constant is 56
   * </p>
   */
  xlAllDatesInPeriodQuarter4(56),
  /**
   * <p>
   * The value of this constant is 57
   * </p>
   */
  xlAllDatesInPeriodJanuary(57),
  /**
   * <p>
   * The value of this constant is 58
   * </p>
   */
  xlAllDatesInPeriodFebruary(58),
  /**
   * <p>
   * The value of this constant is 59
   * </p>
   */
  xlAllDatesInPeriodMarch(59),
  /**
   * <p>
   * The value of this constant is 60
   * </p>
   */
  xlAllDatesInPeriodApril(60),
  /**
   * <p>
   * The value of this constant is 61
   * </p>
   */
  xlAllDatesInPeriodMay(61),
  /**
   * <p>
   * The value of this constant is 62
   * </p>
   */
  xlAllDatesInPeriodJune(62),
  /**
   * <p>
   * The value of this constant is 63
   * </p>
   */
  xlAllDatesInPeriodJuly(63),
  /**
   * <p>
   * The value of this constant is 64
   * </p>
   */
  xlAllDatesInPeriodAugust(64),
  /**
   * <p>
   * The value of this constant is 65
   * </p>
   */
  xlAllDatesInPeriodSeptember(65),
  /**
   * <p>
   * The value of this constant is 66
   * </p>
   */
  xlAllDatesInPeriodOctober(66),
  /**
   * <p>
   * The value of this constant is 67
   * </p>
   */
  xlAllDatesInPeriodNovember(67),
  /**
   * <p>
   * The value of this constant is 68
   * </p>
   */
  xlAllDatesInPeriodDecember(68),
  ;

  private final int value;
  XlPivotFilterType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
