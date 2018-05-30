package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoCondition implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoConditionFileTypeAllFiles(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoConditionFileTypeOfficeFiles(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoConditionFileTypeWordDocuments(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoConditionFileTypeExcelWorkbooks(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoConditionFileTypePowerPointPresentations(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoConditionFileTypeBinders(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  msoConditionFileTypeDatabases(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  msoConditionFileTypeTemplates(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  msoConditionIncludes(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  msoConditionIncludesPhrase(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  msoConditionBeginsWith(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  msoConditionEndsWith(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  msoConditionIncludesNearEachOther(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  msoConditionIsExactly(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  msoConditionIsNot(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  msoConditionYesterday(16),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  msoConditionToday(17),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  msoConditionTomorrow(18),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  msoConditionLastWeek(19),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  msoConditionThisWeek(20),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  msoConditionNextWeek(21),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  msoConditionLastMonth(22),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  msoConditionThisMonth(23),
  /**
   * <p>
   * The value of this constant is 24
   * </p>
   */
  msoConditionNextMonth(24),
  /**
   * <p>
   * The value of this constant is 25
   * </p>
   */
  msoConditionAnytime(25),
  /**
   * <p>
   * The value of this constant is 26
   * </p>
   */
  msoConditionAnytimeBetween(26),
  /**
   * <p>
   * The value of this constant is 27
   * </p>
   */
  msoConditionOn(27),
  /**
   * <p>
   * The value of this constant is 28
   * </p>
   */
  msoConditionOnOrAfter(28),
  /**
   * <p>
   * The value of this constant is 29
   * </p>
   */
  msoConditionOnOrBefore(29),
  /**
   * <p>
   * The value of this constant is 30
   * </p>
   */
  msoConditionInTheNext(30),
  /**
   * <p>
   * The value of this constant is 31
   * </p>
   */
  msoConditionInTheLast(31),
  /**
   * <p>
   * The value of this constant is 32
   * </p>
   */
  msoConditionEquals(32),
  /**
   * <p>
   * The value of this constant is 33
   * </p>
   */
  msoConditionDoesNotEqual(33),
  /**
   * <p>
   * The value of this constant is 34
   * </p>
   */
  msoConditionAnyNumberBetween(34),
  /**
   * <p>
   * The value of this constant is 35
   * </p>
   */
  msoConditionAtMost(35),
  /**
   * <p>
   * The value of this constant is 36
   * </p>
   */
  msoConditionAtLeast(36),
  /**
   * <p>
   * The value of this constant is 37
   * </p>
   */
  msoConditionMoreThan(37),
  /**
   * <p>
   * The value of this constant is 38
   * </p>
   */
  msoConditionLessThan(38),
  /**
   * <p>
   * The value of this constant is 39
   * </p>
   */
  msoConditionIsYes(39),
  /**
   * <p>
   * The value of this constant is 40
   * </p>
   */
  msoConditionIsNo(40),
  /**
   * <p>
   * The value of this constant is 41
   * </p>
   */
  msoConditionIncludesFormsOf(41),
  /**
   * <p>
   * The value of this constant is 42
   * </p>
   */
  msoConditionFreeText(42),
  /**
   * <p>
   * The value of this constant is 43
   * </p>
   */
  msoConditionFileTypeOutlookItems(43),
  /**
   * <p>
   * The value of this constant is 44
   * </p>
   */
  msoConditionFileTypeMailItem(44),
  /**
   * <p>
   * The value of this constant is 45
   * </p>
   */
  msoConditionFileTypeCalendarItem(45),
  /**
   * <p>
   * The value of this constant is 46
   * </p>
   */
  msoConditionFileTypeContactItem(46),
  /**
   * <p>
   * The value of this constant is 47
   * </p>
   */
  msoConditionFileTypeNoteItem(47),
  /**
   * <p>
   * The value of this constant is 48
   * </p>
   */
  msoConditionFileTypeJournalItem(48),
  /**
   * <p>
   * The value of this constant is 49
   * </p>
   */
  msoConditionFileTypeTaskItem(49),
  /**
   * <p>
   * The value of this constant is 50
   * </p>
   */
  msoConditionFileTypePhotoDrawFiles(50),
  /**
   * <p>
   * The value of this constant is 51
   * </p>
   */
  msoConditionFileTypeDataConnectionFiles(51),
  /**
   * <p>
   * The value of this constant is 52
   * </p>
   */
  msoConditionFileTypePublisherFiles(52),
  /**
   * <p>
   * The value of this constant is 53
   * </p>
   */
  msoConditionFileTypeProjectFiles(53),
  /**
   * <p>
   * The value of this constant is 54
   * </p>
   */
  msoConditionFileTypeDocumentImagingFiles(54),
  /**
   * <p>
   * The value of this constant is 55
   * </p>
   */
  msoConditionFileTypeVisioFiles(55),
  /**
   * <p>
   * The value of this constant is 56
   * </p>
   */
  msoConditionFileTypeDesignerFiles(56),
  /**
   * <p>
   * The value of this constant is 57
   * </p>
   */
  msoConditionFileTypeWebPages(57),
  /**
   * <p>
   * The value of this constant is 58
   * </p>
   */
  msoConditionEqualsLow(58),
  /**
   * <p>
   * The value of this constant is 59
   * </p>
   */
  msoConditionEqualsNormal(59),
  /**
   * <p>
   * The value of this constant is 60
   * </p>
   */
  msoConditionEqualsHigh(60),
  /**
   * <p>
   * The value of this constant is 61
   * </p>
   */
  msoConditionNotEqualToLow(61),
  /**
   * <p>
   * The value of this constant is 62
   * </p>
   */
  msoConditionNotEqualToNormal(62),
  /**
   * <p>
   * The value of this constant is 63
   * </p>
   */
  msoConditionNotEqualToHigh(63),
  /**
   * <p>
   * The value of this constant is 64
   * </p>
   */
  msoConditionEqualsNotStarted(64),
  /**
   * <p>
   * The value of this constant is 65
   * </p>
   */
  msoConditionEqualsInProgress(65),
  /**
   * <p>
   * The value of this constant is 66
   * </p>
   */
  msoConditionEqualsCompleted(66),
  /**
   * <p>
   * The value of this constant is 67
   * </p>
   */
  msoConditionEqualsWaitingForSomeoneElse(67),
  /**
   * <p>
   * The value of this constant is 68
   * </p>
   */
  msoConditionEqualsDeferred(68),
  /**
   * <p>
   * The value of this constant is 69
   * </p>
   */
  msoConditionNotEqualToNotStarted(69),
  /**
   * <p>
   * The value of this constant is 70
   * </p>
   */
  msoConditionNotEqualToInProgress(70),
  /**
   * <p>
   * The value of this constant is 71
   * </p>
   */
  msoConditionNotEqualToCompleted(71),
  /**
   * <p>
   * The value of this constant is 72
   * </p>
   */
  msoConditionNotEqualToWaitingForSomeoneElse(72),
  /**
   * <p>
   * The value of this constant is 73
   * </p>
   */
  msoConditionNotEqualToDeferred(73),
  ;

  private final int value;
  MsoCondition(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
