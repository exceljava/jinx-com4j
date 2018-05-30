package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoAppLanguageID implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoLanguageIDInstall(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoLanguageIDUI(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoLanguageIDHelp(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoLanguageIDExeMode(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoLanguageIDUIPrevious(5),
  ;

  private final int value;
  MsoAppLanguageID(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
