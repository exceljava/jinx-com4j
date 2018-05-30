package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSaveConflictResolution implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlLocalSessionChanges(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlOtherSessionChanges(3),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlUserResolution(1),
  ;

  private final int value;
  XlSaveConflictResolution(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
