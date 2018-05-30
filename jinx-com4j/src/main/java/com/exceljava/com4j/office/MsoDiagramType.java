package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoDiagramType implements ComEnum {
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoDiagramMixed(-2),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoDiagramOrgChart(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoDiagramCycle(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoDiagramRadial(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoDiagramPyramid(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoDiagramVenn(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoDiagramTarget(6),
  ;

  private final int value;
  MsoDiagramType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
