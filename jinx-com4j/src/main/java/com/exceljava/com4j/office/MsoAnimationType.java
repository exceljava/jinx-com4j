package com.exceljava.com4j.office  ;

import com4j.*;

/**
 */
public enum MsoAnimationType implements ComEnum {
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  msoAnimationIdle(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  msoAnimationGreeting(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  msoAnimationGoodbye(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  msoAnimationBeginSpeaking(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  msoAnimationRestPose(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  msoAnimationCharacterSuccessMajor(6),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  msoAnimationGetAttentionMajor(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  msoAnimationGetAttentionMinor(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  msoAnimationSearching(13),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  msoAnimationPrinting(18),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  msoAnimationGestureRight(19),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  msoAnimationWritingNotingSomething(22),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  msoAnimationWorkingAtSomething(23),
  /**
   * <p>
   * The value of this constant is 24
   * </p>
   */
  msoAnimationThinking(24),
  /**
   * <p>
   * The value of this constant is 25
   * </p>
   */
  msoAnimationSendingMail(25),
  /**
   * <p>
   * The value of this constant is 26
   * </p>
   */
  msoAnimationListensToComputer(26),
  /**
   * <p>
   * The value of this constant is 31
   * </p>
   */
  msoAnimationDisappear(31),
  /**
   * <p>
   * The value of this constant is 32
   * </p>
   */
  msoAnimationAppear(32),
  /**
   * <p>
   * The value of this constant is 100
   * </p>
   */
  msoAnimationGetArtsy(100),
  /**
   * <p>
   * The value of this constant is 101
   * </p>
   */
  msoAnimationGetTechy(101),
  /**
   * <p>
   * The value of this constant is 102
   * </p>
   */
  msoAnimationGetWizardy(102),
  /**
   * <p>
   * The value of this constant is 103
   * </p>
   */
  msoAnimationCheckingSomething(103),
  /**
   * <p>
   * The value of this constant is 104
   * </p>
   */
  msoAnimationLookDown(104),
  /**
   * <p>
   * The value of this constant is 105
   * </p>
   */
  msoAnimationLookDownLeft(105),
  /**
   * <p>
   * The value of this constant is 106
   * </p>
   */
  msoAnimationLookDownRight(106),
  /**
   * <p>
   * The value of this constant is 107
   * </p>
   */
  msoAnimationLookLeft(107),
  /**
   * <p>
   * The value of this constant is 108
   * </p>
   */
  msoAnimationLookRight(108),
  /**
   * <p>
   * The value of this constant is 109
   * </p>
   */
  msoAnimationLookUp(109),
  /**
   * <p>
   * The value of this constant is 110
   * </p>
   */
  msoAnimationLookUpLeft(110),
  /**
   * <p>
   * The value of this constant is 111
   * </p>
   */
  msoAnimationLookUpRight(111),
  /**
   * <p>
   * The value of this constant is 112
   * </p>
   */
  msoAnimationSaving(112),
  /**
   * <p>
   * The value of this constant is 113
   * </p>
   */
  msoAnimationGestureDown(113),
  /**
   * <p>
   * The value of this constant is 114
   * </p>
   */
  msoAnimationGestureLeft(114),
  /**
   * <p>
   * The value of this constant is 115
   * </p>
   */
  msoAnimationGestureUp(115),
  /**
   * <p>
   * The value of this constant is 116
   * </p>
   */
  msoAnimationEmptyTrash(116),
  ;

  private final int value;
  MsoAnimationType(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
