package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPictureConvertorType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlBMP(1),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlCGM(7),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlDRW(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlDXF(5),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlEPS(8),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlHGL(6),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  xlPCT(13),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlPCX(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlPIC(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  xlPLT(12),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlTIF(9),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlWMF(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlWPG(3),
  ;

  private final int value;
  XlPictureConvertorType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
