package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoMoveRow implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4
   * </p>
   */
  msoMoveRowFirst(-4),
  /**
   * <p>
   * The value of this constant is -3
   * </p>
   */
  msoMoveRowPrev(-3),
  /**
   * <p>
   * The value of this constant is -2
   * </p>
   */
  msoMoveRowNext(-2),
  /**
   * <p>
   * The value of this constant is -1
   * </p>
   */
  msoMoveRowNbr(-1),
  ;

  private final int value;
  MsoMoveRow(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
