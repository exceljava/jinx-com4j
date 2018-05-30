package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoReflectionType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoReflectionTypeMixed(-2),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoReflectionTypeNone(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoReflectionType1(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoReflectionType2(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoReflectionType3(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoReflectionType4(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoReflectionType5(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoReflectionType6(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoReflectionType7(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoReflectionType8(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoReflectionType9(9),
  ;

  private final int value;
  MsoReflectionType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
