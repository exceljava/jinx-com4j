package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoFontLanguageIndex implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoThemeLatin(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoThemeComplexScript(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoThemeEastAsian(3),
  ;

  private final int value;
  MsoFontLanguageIndex(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
