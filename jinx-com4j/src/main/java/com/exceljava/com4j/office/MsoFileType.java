package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoFileType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoFileTypeAllFiles(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoFileTypeOfficeFiles(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoFileTypeWordDocuments(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoFileTypeExcelWorkbooks(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoFileTypePowerPointPresentations(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoFileTypeBinders(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoFileTypeDatabases(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoFileTypeTemplates(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoFileTypeOutlookItems(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  msoFileTypeMailItem(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  msoFileTypeCalendarItem(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  msoFileTypeContactItem(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  msoFileTypeNoteItem(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  msoFileTypeJournalItem(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  msoFileTypeTaskItem(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  msoFileTypePhotoDrawFiles(16),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  msoFileTypeDataConnectionFiles(17),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  msoFileTypePublisherFiles(18),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  msoFileTypeProjectFiles(19),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  msoFileTypeDocumentImagingFiles(20),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  msoFileTypeVisioFiles(21),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  msoFileTypeDesignerFiles(22),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  msoFileTypeWebPages(23),
  ;

  private final int value;
  MsoFileType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
