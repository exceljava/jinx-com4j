package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoFileFindSortBy implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoFileFindSortbyAuthor(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoFileFindSortbyDateCreated(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoFileFindSortbyLastSavedBy(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoFileFindSortbyDateSaved(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoFileFindSortbyFileName(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoFileFindSortbySize(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoFileFindSortbyTitle(7),
  ;

  private final int value;
  MsoFileFindSortBy(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
