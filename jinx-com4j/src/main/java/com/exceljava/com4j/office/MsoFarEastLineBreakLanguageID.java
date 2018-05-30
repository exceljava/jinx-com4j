package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoFarEastLineBreakLanguageID implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1041
   * </p>
   */
  MsoFarEastLineBreakLanguageJapanese(1041),
  /**
   * <p>
   * The value of this constant is 1042
   * </p>
   */
  MsoFarEastLineBreakLanguageKorean(1042),
  /**
   * <p>
   * The value of this constant is 2052
   * </p>
   */
  MsoFarEastLineBreakLanguageSimplifiedChinese(2052),
  /**
   * <p>
   * The value of this constant is 1028
   * </p>
   */
  MsoFarEastLineBreakLanguageTraditionalChinese(1028),
  ;

  private final int value;
  MsoFarEastLineBreakLanguageID(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
