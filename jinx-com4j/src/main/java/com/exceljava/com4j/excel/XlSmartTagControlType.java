package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlSmartTagControlType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlSmartTagControlSmartTag(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlSmartTagControlLink(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlSmartTagControlHelp(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlSmartTagControlHelpURL(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlSmartTagControlSeparator(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlSmartTagControlButton(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlSmartTagControlLabel(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlSmartTagControlImage(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlSmartTagControlCheckbox(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlSmartTagControlTextbox(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlSmartTagControlListbox(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  xlSmartTagControlCombo(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  xlSmartTagControlActiveX(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  xlSmartTagControlRadioGroup(14),
  ;

  private final int value;
  XlSmartTagControlType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
