package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlYesNoGuess implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  xlGuess(0),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlNo(2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlYes(1),
  ;

  private final int value;
  XlYesNoGuess(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
