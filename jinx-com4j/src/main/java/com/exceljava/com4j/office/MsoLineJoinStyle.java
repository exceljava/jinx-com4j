package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoLineJoinStyle implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoLineJoinMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoLineJoinRound(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoLineJoinBevel(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoLineJoinMiter(3),
  ;

  private final int value;
  MsoLineJoinStyle(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
