package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoBevelType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoBevelTypeMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoBevelNone(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoBevelRelaxedInset(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoBevelCircle(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoBevelSlope(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoBevelCross(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoBevelAngle(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoBevelSoftRound(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoBevelConvex(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoBevelCoolSlant(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  msoBevelDivot(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  msoBevelRiblet(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  msoBevelHardEdge(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  msoBevelArtDeco(13),
  ;

  private final int value;
  MsoBevelType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
