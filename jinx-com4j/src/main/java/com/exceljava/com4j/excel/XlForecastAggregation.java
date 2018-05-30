package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlForecastAggregation implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlForecastAggregationAverage(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlForecastAggregationCount(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlForecastAggregationCountA(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlForecastAggregationMax(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlForecastAggregationMedian(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlForecastAggregationMin(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlForecastAggregationSum(7),
  ;

  private final int value;
  XlForecastAggregation(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
