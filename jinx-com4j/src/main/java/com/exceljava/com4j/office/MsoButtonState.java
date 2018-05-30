package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoButtonState implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoButtonUp(0),
  /**
   * <p>
   * The value of this constant is -1
   * </p>
   */
  msoButtonDown(-1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoButtonMixed(2),
  ;

  private final int value;
  MsoButtonState(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
