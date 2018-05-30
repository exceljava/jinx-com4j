package com.exceljava.com4j.vbide  ;

import com4j.*;

/**
 */
public enum vbext_WindowType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  vbext_wt_CodeWindow(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  vbext_wt_Designer(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  vbext_wt_Browser(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  vbext_wt_Watch(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  vbext_wt_Locals(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  vbext_wt_Immediate(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  vbext_wt_ProjectWindow(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  vbext_wt_PropertyWindow(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  vbext_wt_Find(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  vbext_wt_FindReplace(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  vbext_wt_Toolbox(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  vbext_wt_LinkedWindowFrame(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  vbext_wt_MainWindow(12),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  vbext_wt_ToolWindow(15),
  ;

  private final int value;
  vbext_WindowType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
