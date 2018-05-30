package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlCubeFieldType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlHierarchy(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlMeasure(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlSet(3),
  ;

  private final int value;
  XlCubeFieldType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
