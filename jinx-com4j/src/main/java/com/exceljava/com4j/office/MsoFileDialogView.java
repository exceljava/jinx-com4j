package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoFileDialogView implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoFileDialogViewList(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoFileDialogViewDetails(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoFileDialogViewProperties(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoFileDialogViewPreview(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoFileDialogViewThumbnail(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoFileDialogViewLargeIcons(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoFileDialogViewSmallIcons(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoFileDialogViewWebView(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoFileDialogViewTiles(9),
  ;

  private final int value;
  MsoFileDialogView(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
