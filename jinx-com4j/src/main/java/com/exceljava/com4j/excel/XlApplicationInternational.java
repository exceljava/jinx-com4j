package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlApplicationInternational implements ComEnum {
  /**
   * <p>
   * The value of this constant is 33
   * </p>
   */
  xl24HourClock(33),
  /**
   * <p>
   * The value of this constant is 43
   * </p>
   */
  xl4DigitYears(43),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  xlAlternateArraySeparator(16),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  xlColumnSeparator(14),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlCountryCode(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlCountrySetting(2),
  /**
   * <p>
   * The value of this constant is 37
   * </p>
   */
  xlCurrencyBefore(37),
  /**
   * <p>
   * The value of this constant is 25
   * </p>
   */
  xlCurrencyCode(25),
  /**
   * <p>
   * The value of this constant is 27
   * </p>
   */
  xlCurrencyDigits(27),
  /**
   * <p>
   * The value of this constant is 40
   * </p>
   */
  xlCurrencyLeadingZeros(40),
  /**
   * <p>
   * The value of this constant is 38
   * </p>
   */
  xlCurrencyMinusSign(38),
  /**
   * <p>
   * The value of this constant is 28
   * </p>
   */
  xlCurrencyNegative(28),
  /**
   * <p>
   * The value of this constant is 36
   * </p>
   */
  xlCurrencySpaceBefore(36),
  /**
   * <p>
   * The value of this constant is 39
   * </p>
   */
  xlCurrencyTrailingZeros(39),
  /**
   * <p>
   * The value of this constant is 32
   * </p>
   */
  xlDateOrder(32),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  xlDateSeparator(17),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  xlDayCode(21),
  /**
   * <p>
   * The value of this constant is 42
   * </p>
   */
  xlDayLeadingZero(42),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlDecimalSeparator(3),
  /**
   * <p>
   * The value of this constant is 26
   * </p>
   */
  xlGeneralFormatName(26),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  xlHourCode(22),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  xlLeftBrace(12),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlLeftBracket(10),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlListSeparator(5),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlLowerCaseColumnLetter(9),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlLowerCaseRowLetter(8),
  /**
   * <p>
   * The value of this constant is 44
   * </p>
   */
  xlMDY(44),
  /**
   * <p>
   * The value of this constant is 35
   * </p>
   */
  xlMetric(35),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  xlMinuteCode(23),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  xlMonthCode(20),
  /**
   * <p>
   * The value of this constant is 41
   * </p>
   */
  xlMonthLeadingZero(41),
  /**
   * <p>
   * The value of this constant is 30
   * </p>
   */
  xlMonthNameChars(30),
  /**
   * <p>
   * The value of this constant is 29
   * </p>
   */
  xlNoncurrencyDigits(29),
  /**
   * <p>
   * The value of this constant is 34
   * </p>
   */
  xlNonEnglishFunctions(34),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  xlRightBrace(13),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlRightBracket(11),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  xlRowSeparator(15),
  /**
   * <p>
   * The value of this constant is 24
   * </p>
   */
  xlSecondCode(24),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlThousandsSeparator(4),
  /**
   * <p>
   * The value of this constant is 45
   * </p>
   */
  xlTimeLeadingZero(45),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  xlTimeSeparator(18),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlUpperCaseColumnLetter(7),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlUpperCaseRowLetter(6),
  /**
   * <p>
   * The value of this constant is 31
   * </p>
   */
  xlWeekdayNameChars(31),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  xlYearCode(19),
  /**
   * <p>
   * The value of this constant is 46
   * </p>
   */
  xlUICultureTag(46),
  ;

  private final int value;
  XlApplicationInternational(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
