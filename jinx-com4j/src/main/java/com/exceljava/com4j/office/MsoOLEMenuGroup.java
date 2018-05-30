package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoOLEMenuGroup implements ComEnum {
  /**
   * <p>
   * The value of this constant is -1
   * </p>
   */
  msoOLEMenuGroupNone(-1),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  msoOLEMenuGroupFile(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoOLEMenuGroupEdit(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoOLEMenuGroupContainer(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoOLEMenuGroupObject(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoOLEMenuGroupWindow(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoOLEMenuGroupHelp(5),
  ;

  private final int value;
  MsoOLEMenuGroup(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
