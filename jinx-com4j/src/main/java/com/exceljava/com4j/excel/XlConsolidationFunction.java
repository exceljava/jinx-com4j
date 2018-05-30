package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlConsolidationFunction implements ComEnum {
  /**
   * <p>
   * The value of this constant is -4106
   * </p>
   */
  xlAverage(-4106),
  /**
   * <p>
   * The value of this constant is -4112
   * </p>
   */
  xlCount(-4112),
  /**
   * <p>
   * The value of this constant is -4113
   * </p>
   */
  xlCountNums(-4113),
  /**
   * <p>
   * The value of this constant is -4136
   * </p>
   */
  xlMax(-4136),
  /**
   * <p>
   * The value of this constant is -4139
   * </p>
   */
  xlMin(-4139),
  /**
   * <p>
   * The value of this constant is -4149
   * </p>
   */
  xlProduct(-4149),
  /**
   * <p>
   * The value of this constant is -4155
   * </p>
   */
  xlStDev(-4155),
  /**
   * <p>
   * The value of this constant is -4156
   * </p>
   */
  xlStDevP(-4156),
  /**
   * <p>
   * The value of this constant is -4157
   * </p>
   */
  xlSum(-4157),
  /**
   * <p>
   * The value of this constant is -4164
   * </p>
   */
  xlVar(-4164),
  /**
   * <p>
   * The value of this constant is -4165
   * </p>
   */
  xlVarP(-4165),
  /**
   * <p>
   * The value of this constant is 1000
   * </p>
   */
  xlUnknown(1000),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlDistinctCount(11),
  ;

  private final int value;
  XlConsolidationFunction(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
