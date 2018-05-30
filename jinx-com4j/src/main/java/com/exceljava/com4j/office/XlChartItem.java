package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum XlChartItem implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  xlDataLabel(0),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlChartArea(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlSeries(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlChartTitle(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlWalls(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlCorners(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlDataTable(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlTrendline(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlErrorBars(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlXErrorBars(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlYErrorBars(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  xlLegendEntry(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  xlLegendKey(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  xlShape(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  xlMajorGridlines(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  xlMinorGridlines(16),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  xlAxisTitle(17),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  xlUpBars(18),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  xlPlotArea(19),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  xlDownBars(20),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  xlAxis(21),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  xlSeriesLines(22),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  xlFloor(23),
  /**
   * <p>
   * The value of this constant is 24
   * </p>
   */
  xlLegend(24),
  /**
   * <p>
   * The value of this constant is 25
   * </p>
   */
  xlHiLoLines(25),
  /**
   * <p>
   * The value of this constant is 26
   * </p>
   */
  xlDropLines(26),
  /**
   * <p>
   * The value of this constant is 27
   * </p>
   */
  xlRadarAxisLabels(27),
  /**
   * <p>
   * The value of this constant is 28
   * </p>
   */
  xlNothing(28),
  /**
   * <p>
   * The value of this constant is 29
   * </p>
   */
  xlLeaderLines(29),
  /**
   * <p>
   * The value of this constant is 30
   * </p>
   */
  xlDisplayUnitLabel(30),
  /**
   * <p>
   * The value of this constant is 31
   * </p>
   */
  xlPivotChartFieldButton(31),
  /**
   * <p>
   * The value of this constant is 32
   * </p>
   */
  xlPivotChartDropZone(32),
  /**
   * <p>
   * The value of this constant is 33
   * </p>
   */
  xlPivotChartExpandEntireFieldButton(33),
  /**
   * <p>
   * The value of this constant is 34
   * </p>
   */
  xlPivotChartCollapseEntireFieldButton(34),
  ;

  private final int value;
  XlChartItem(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
