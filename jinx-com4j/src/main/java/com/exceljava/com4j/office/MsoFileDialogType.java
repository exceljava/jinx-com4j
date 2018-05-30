package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoFileDialogType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoFileDialogOpen(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoFileDialogSaveAs(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoFileDialogFilePicker(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoFileDialogFolderPicker(4),
  ;

  private final int value;
  MsoFileDialogType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
