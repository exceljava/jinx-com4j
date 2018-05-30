package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlIconSet implements ComEnum {
  /**
   * <p>
   * The value of this constant is -1
   * </p>
   */
  xlCustomSet(-1),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xl3Arrows(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xl3ArrowsGray(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xl3Flags(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xl3TrafficLights1(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xl3TrafficLights2(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xl3Signs(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xl3Symbols(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xl3Symbols2(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xl4Arrows(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xl4ArrowsGray(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xl4RedToBlack(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  xl4CRV(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  xl4TrafficLights(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  xl5Arrows(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  xl5ArrowsGray(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  xl5CRV(16),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  xl5Quarters(17),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  xl3Stars(18),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  xl3Triangles(19),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  xl5Boxes(20),
  ;

  private final int value;
  XlIconSet(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
