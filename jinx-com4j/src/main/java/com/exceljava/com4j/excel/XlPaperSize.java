package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlPaperSize implements ComEnum {
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  xlPaper10x14(16),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  xlPaper11x17(17),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlPaperA3(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlPaperA4(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlPaperA4Small(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlPaperA5(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  xlPaperB4(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  xlPaperB5(13),
  /**
   * <p>
   * The value of this constant is 24
   * </p>
   */
  xlPaperCsheet(24),
  /**
   * <p>
   * The value of this constant is 25
   * </p>
   */
  xlPaperDsheet(25),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  xlPaperEnvelope10(20),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  xlPaperEnvelope11(21),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  xlPaperEnvelope12(22),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  xlPaperEnvelope14(23),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  xlPaperEnvelope9(19),
  /**
   * <p>
   * The value of this constant is 33
   * </p>
   */
  xlPaperEnvelopeB4(33),
  /**
   * <p>
   * The value of this constant is 34
   * </p>
   */
  xlPaperEnvelopeB5(34),
  /**
   * <p>
   * The value of this constant is 35
   * </p>
   */
  xlPaperEnvelopeB6(35),
  /**
   * <p>
   * The value of this constant is 29
   * </p>
   */
  xlPaperEnvelopeC3(29),
  /**
   * <p>
   * The value of this constant is 30
   * </p>
   */
  xlPaperEnvelopeC4(30),
  /**
   * <p>
   * The value of this constant is 28
   * </p>
   */
  xlPaperEnvelopeC5(28),
  /**
   * <p>
   * The value of this constant is 31
   * </p>
   */
  xlPaperEnvelopeC6(31),
  /**
   * <p>
   * The value of this constant is 32
   * </p>
   */
  xlPaperEnvelopeC65(32),
  /**
   * <p>
   * The value of this constant is 27
   * </p>
   */
  xlPaperEnvelopeDL(27),
  /**
   * <p>
   * The value of this constant is 36
   * </p>
   */
  xlPaperEnvelopeItaly(36),
  /**
   * <p>
   * The value of this constant is 37
   * </p>
   */
  xlPaperEnvelopeMonarch(37),
  /**
   * <p>
   * The value of this constant is 38
   * </p>
   */
  xlPaperEnvelopePersonal(38),
  /**
   * <p>
   * The value of this constant is 26
   * </p>
   */
  xlPaperEsheet(26),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlPaperExecutive(7),
  /**
   * <p>
   * The value of this constant is 41
   * </p>
   */
  xlPaperFanfoldLegalGerman(41),
  /**
   * <p>
   * The value of this constant is 40
   * </p>
   */
  xlPaperFanfoldStdGerman(40),
  /**
   * <p>
   * The value of this constant is 39
   * </p>
   */
  xlPaperFanfoldUS(39),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  xlPaperFolio(14),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlPaperLedger(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlPaperLegal(5),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlPaperLetter(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlPaperLetterSmall(2),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  xlPaperNote(18),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  xlPaperQuarto(15),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlPaperStatement(6),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlPaperTabloid(3),
  /**
   * <p>
   * The value of this constant is 256
   * </p>
   */
  xlPaperUser(256),
  ;

  private final int value;
  XlPaperSize(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
