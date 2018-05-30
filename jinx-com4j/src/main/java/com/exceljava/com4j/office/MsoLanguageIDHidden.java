package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoLanguageIDHidden implements ComEnum {
  /**
   * <p>
   * The value of this constant is 3076
   * </p>
   */
  msoLanguageIDChineseHongKong(3076),
  /**
   * <p>
   * The value of this constant is 5124
   * </p>
   */
  msoLanguageIDChineseMacao(5124),
  /**
   * <p>
   * The value of this constant is 11273
   * </p>
   */
  msoLanguageIDEnglishTrinidad(11273),
  ;

  private final int value;
  MsoLanguageIDHidden(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
