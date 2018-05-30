package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPivotTableSourceType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlScenario(4),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlConsolidation(3),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlDatabase(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlExternal(2),
  /**
   * <p>
   * The value of this constant is -4148
   * </p>
   */
  xlPivotTable(-4148),
  ;

  private final int value;
  XlPivotTableSourceType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
