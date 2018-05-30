package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoConnectorType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoConnectorTypeMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoConnectorStraight(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoConnectorElbow(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoConnectorCurve(3),
  ;

  private final int value;
  MsoConnectorType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
