package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlToolbarProtection implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlNoButtonChanges(1),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlNoChanges(4),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlNoDockingChanges(3),
  /**
   * <p>
   * The value of this constant is -4143
   * </p>
   */
  xlToolbarProtectionNone(-4143),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlNoShapeChanges(2),
  ;

  private final int value;
  XlToolbarProtection(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
