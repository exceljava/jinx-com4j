package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoHTMLProjectOpen implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoHTMLProjectOpenSourceView(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoHTMLProjectOpenTextView(2),
  ;

  private final int value;
  MsoHTMLProjectOpen(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
