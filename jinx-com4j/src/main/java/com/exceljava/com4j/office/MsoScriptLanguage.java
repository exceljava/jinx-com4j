package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoScriptLanguage implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoScriptLanguageJava(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoScriptLanguageVisualBasic(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoScriptLanguageASP(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoScriptLanguageOther(4),
  ;

  private final int value;
  MsoScriptLanguage(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
