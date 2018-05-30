package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlTableStyleElementType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  xlWholeTable(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlHeaderRow(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlTotalRow(2),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlGrandTotalRow(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlFirstColumn(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlLastColumn(4),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlGrandTotalColumn(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlRowStripe1(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlRowStripe2(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlColumnStripe1(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlColumnStripe2(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlFirstHeaderCell(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlLastHeaderCell(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlFirstTotalCell(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  xlLastTotalCell(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  xlSubtotalColumn1(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  xlSubtotalColumn2(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  xlSubtotalColumn3(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  xlSubtotalRow1(16),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  xlSubtotalRow2(17),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  xlSubtotalRow3(18),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  xlBlankRow(19),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  xlColumnSubheading1(20),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  xlColumnSubheading2(21),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  xlColumnSubheading3(22),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  xlRowSubheading1(23),
  /**
   * <p>
   * The value of this constant is 24
   * </p>
   */
  xlRowSubheading2(24),
  /**
   * <p>
   * The value of this constant is 25
   * </p>
   */
  xlRowSubheading3(25),
  /**
   * <p>
   * The value of this constant is 26
   * </p>
   */
  xlPageFieldLabels(26),
  /**
   * <p>
   * The value of this constant is 27
   * </p>
   */
  xlPageFieldValues(27),
  /**
   * <p>
   * The value of this constant is 28
   * </p>
   */
  xlSlicerUnselectedItemWithData(28),
  /**
   * <p>
   * The value of this constant is 29
   * </p>
   */
  xlSlicerUnselectedItemWithNoData(29),
  /**
   * <p>
   * The value of this constant is 30
   * </p>
   */
  xlSlicerSelectedItemWithData(30),
  /**
   * <p>
   * The value of this constant is 31
   * </p>
   */
  xlSlicerSelectedItemWithNoData(31),
  /**
   * <p>
   * The value of this constant is 32
   * </p>
   */
  xlSlicerHoveredUnselectedItemWithData(32),
  /**
   * <p>
   * The value of this constant is 33
   * </p>
   */
  xlSlicerHoveredSelectedItemWithData(33),
  /**
   * <p>
   * The value of this constant is 34
   * </p>
   */
  xlSlicerHoveredUnselectedItemWithNoData(34),
  /**
   * <p>
   * The value of this constant is 35
   * </p>
   */
  xlSlicerHoveredSelectedItemWithNoData(35),
  /**
   * <p>
   * The value of this constant is 36
   * </p>
   */
  xlTimelineSelectionLabel(36),
  /**
   * <p>
   * The value of this constant is 37
   * </p>
   */
  xlTimelineTimeLevel(37),
  /**
   * <p>
   * The value of this constant is 38
   * </p>
   */
  xlTimelinePeriodLabels1(38),
  /**
   * <p>
   * The value of this constant is 39
   * </p>
   */
  xlTimelinePeriodLabels2(39),
  /**
   * <p>
   * The value of this constant is 40
   * </p>
   */
  xlTimelineSelectedTimeBlock(40),
  /**
   * <p>
   * The value of this constant is 41
   * </p>
   */
  xlTimelineUnselectedTimeBlock(41),
  /**
   * <p>
   * The value of this constant is 42
   * </p>
   */
  xlTimelineSelectedTimeBlockSpace(42),
  ;

  private final int value;
  XlTableStyleElementType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
