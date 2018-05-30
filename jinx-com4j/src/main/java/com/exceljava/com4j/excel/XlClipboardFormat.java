package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlClipboardFormat implements ComEnum {
  /**
   * <p>
   * The value of this constant is 63
   * </p>
   */
  xlClipboardFormatBIFF12(63),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlClipboardFormatBIFF(8),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  xlClipboardFormatBIFF2(18),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  xlClipboardFormatBIFF3(20),
  /**
   * <p>
   * The value of this constant is 30
   * </p>
   */
  xlClipboardFormatBIFF4(30),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  xlClipboardFormatBinary(15),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlClipboardFormatBitmap(9),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  xlClipboardFormatCGM(13),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlClipboardFormatCSV(5),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlClipboardFormatDIF(4),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  xlClipboardFormatDspText(12),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  xlClipboardFormatEmbeddedObject(21),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  xlClipboardFormatEmbedSource(22),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlClipboardFormatLink(11),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  xlClipboardFormatLinkSource(23),
  /**
   * <p>
   * The value of this constant is 32
   * </p>
   */
  xlClipboardFormatLinkSourceDesc(32),
  /**
   * <p>
   * The value of this constant is 24
   * </p>
   */
  xlClipboardFormatMovie(24),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  xlClipboardFormatNative(14),
  /**
   * <p>
   * The value of this constant is 31
   * </p>
   */
  xlClipboardFormatObjectDesc(31),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  xlClipboardFormatObjectLink(19),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  xlClipboardFormatOwnerLink(17),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlClipboardFormatPICT(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  xlClipboardFormatPrintPICT(3),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlClipboardFormatRTF(7),
  /**
   * <p>
   * The value of this constant is 29
   * </p>
   */
  xlClipboardFormatScreenPICT(29),
  /**
   * <p>
   * The value of this constant is 28
   * </p>
   */
  xlClipboardFormatStandardFont(28),
  /**
   * <p>
   * The value of this constant is 27
   * </p>
   */
  xlClipboardFormatStandardScale(27),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlClipboardFormatSYLK(6),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  xlClipboardFormatTable(16),
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  xlClipboardFormatText(0),
  /**
   * <p>
   * The value of this constant is 25
   * </p>
   */
  xlClipboardFormatToolFace(25),
  /**
   * <p>
   * The value of this constant is 26
   * </p>
   */
  xlClipboardFormatToolFacePICT(26),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  xlClipboardFormatVALU(1),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  xlClipboardFormatWK1(10),
  ;

  private final int value;
  XlClipboardFormat(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
