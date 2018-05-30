package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlCellChangedState implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlCellNotChanged(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlCellChanged(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlCellChangeApplied(3),
  ;

  private final int value;
  XlCellChangedState(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
