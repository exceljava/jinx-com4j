package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoTriState implements ComEnum {
  /**
   * <p>
   * The value of this constant is -1
   * </p>
   */
  msoTrue(-1),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoFalse(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoCTrue(1),
  /**
   * <p>
   * The value of this constant is -3
   * </p>
   */
  msoTriStateToggle(-3),
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoTriStateMixed(-2),
  ;

  private final int value;
  MsoTriState(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
