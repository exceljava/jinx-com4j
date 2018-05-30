package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MailFormat implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  mfPlainText(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  mfHTML(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  mfRTF(3),
  ;

  private final int value;
  MailFormat(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
