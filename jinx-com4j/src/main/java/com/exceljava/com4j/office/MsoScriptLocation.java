package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoScriptLocation implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoScriptLocationInHead(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoScriptLocationInBody(2),
  ;

  private final int value;
  MsoScriptLocation(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
