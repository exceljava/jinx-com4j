package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlMSApplication implements ComEnum {
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlMicrosoftAccess(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlMicrosoftFoxPro(5),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlMicrosoftMail(3),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlMicrosoftPowerPoint(2),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlMicrosoftProject(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlMicrosoftSchedulePlus(7),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlMicrosoftWord(1),
  ;

  private final int value;
  XlMSApplication(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
