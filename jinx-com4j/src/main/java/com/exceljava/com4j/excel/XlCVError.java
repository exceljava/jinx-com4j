package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlCVError implements ComEnum {
  /**
   * <p>
   * The value of this constant is 2047
   * </p>
   */
  xlErrBlocked(2047),
  /**
   * <p>
   * The value of this constant is 2050
   * </p>
   */
  xlErrCalc(2050),
  /**
   * <p>
   * The value of this constant is 2046
   * </p>
   */
  xlErrConnect(2046),
  /**
   * <p>
   * The value of this constant is 2007
   * </p>
   */
  xlErrDiv0(2007),
  /**
   * <p>
   * The value of this constant is 2049
   * </p>
   */
  xlErrField(2049),
  /**
   * <p>
   * The value of this constant is 2043
   * </p>
   */
  xlErrGettingData(2043),
  /**
   * <p>
   * The value of this constant is 2042
   * </p>
   */
  xlErrNA(2042),
  /**
   * <p>
   * The value of this constant is 2029
   * </p>
   */
  xlErrName(2029),
  /**
   * <p>
   * The value of this constant is 2045
   * </p>
   */
  xlErrSpill(2045),
  /**
   * <p>
   * The value of this constant is 2000
   * </p>
   */
  xlErrNull(2000),
  /**
   * <p>
   * The value of this constant is 2036
   * </p>
   */
  xlErrNum(2036),
  /**
   * <p>
   * The value of this constant is 2023
   * </p>
   */
  xlErrRef(2023),
  /**
   * <p>
   * The value of this constant is 2048
   * </p>
   */
  xlErrUnknown(2048),
  /**
   * <p>
   * The value of this constant is 2015
   * </p>
   */
  xlErrValue(2015),
  ;

  private final int value;
  XlCVError(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
