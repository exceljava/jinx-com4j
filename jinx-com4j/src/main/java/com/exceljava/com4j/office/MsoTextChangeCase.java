package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoTextChangeCase implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoCaseSentence(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoCaseLower(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoCaseUpper(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoCaseTitle(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoCaseToggle(5),
  ;

  private final int value;
  MsoTextChangeCase(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
