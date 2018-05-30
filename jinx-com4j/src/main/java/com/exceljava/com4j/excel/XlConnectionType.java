package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlConnectionType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlConnectionTypeOLEDB(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlConnectionTypeODBC(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlConnectionTypeXMLMAP(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlConnectionTypeTEXT(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlConnectionTypeWEB(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlConnectionTypeDATAFEED(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlConnectionTypeMODEL(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlConnectionTypeWORKSHEET(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlConnectionTypeNOSOURCE(9),
  ;

  private final int value;
  XlConnectionType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
