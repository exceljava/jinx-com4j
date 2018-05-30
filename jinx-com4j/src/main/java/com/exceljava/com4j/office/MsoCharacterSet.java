package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoCharacterSet implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoCharacterSetArabic(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoCharacterSetCyrillic(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoCharacterSetEnglishWesternEuropeanOtherLatinScript(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoCharacterSetGreek(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoCharacterSetHebrew(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoCharacterSetJapanese(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoCharacterSetKorean(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoCharacterSetMultilingualUnicode(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoCharacterSetSimplifiedChinese(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  msoCharacterSetThai(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  msoCharacterSetTraditionalChinese(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  msoCharacterSetVietnamese(12),
  ;

  private final int value;
  MsoCharacterSet(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
