package com.exceljava.com4j.excel  ;

import com4j.*;

/**
 */
public enum XlFileFormat implements ComEnum {
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  xlAddIn(18),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  xlCSV(6),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  xlCSVMac(22),
  /**
   * <p>
   * The value of this constant is 24
   * </p>
   */
  xlCSVMSDOS(24),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  xlCSVWindows(23),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  xlDBF2(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  xlDBF3(8),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  xlDBF4(11),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  xlDIF(9),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  xlExcel2(16),
  /**
   * <p>
   * The value of this constant is 27
   * </p>
   */
  xlExcel2FarEast(27),
  /**
   * <p>
   * The value of this constant is 29
   * </p>
   */
  xlExcel3(29),
  /**
   * <p>
   * The value of this constant is 33
   * </p>
   */
  xlExcel4(33),
  /**
   * <p>
   * The value of this constant is 39
   * </p>
   */
  xlExcel5(39),
  /**
   * <p>
   * The value of this constant is 39
   * </p>
   */
  xlExcel7(39),
  /**
   * <p>
   * The value of this constant is 43
   * </p>
   */
  xlExcel9795(43),
  /**
   * <p>
   * The value of this constant is 35
   * </p>
   */
  xlExcel4Workbook(35),
  /**
   * <p>
   * The value of this constant is 26
   * </p>
   */
  xlIntlAddIn(26),
  /**
   * <p>
   * The value of this constant is 25
   * </p>
   */
  xlIntlMacro(25),
  /**
   * <p>
   * The value of this constant is -4143
   * </p>
   */
  xlWorkbookNormal(-4143),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  xlSYLK(2),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  xlTemplate(17),
  /**
   * <p>
   * The value of this constant is -4158
   * </p>
   */
  xlCurrentPlatformText(-4158),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  xlTextMac(19),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  xlTextMSDOS(21),
  /**
   * <p>
   * The value of this constant is 36
   * </p>
   */
  xlTextPrinter(36),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  xlTextWindows(20),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  xlWJ2WD1(14),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  xlWK1(5),
  /**
   * <p>
   * The value of this constant is 31
   * </p>
   */
  xlWK1ALL(31),
  /**
   * <p>
   * The value of this constant is 30
   * </p>
   */
  xlWK1FMT(30),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  xlWK3(15),
  /**
   * <p>
   * The value of this constant is 38
   * </p>
   */
  xlWK4(38),
  /**
   * <p>
   * The value of this constant is 32
   * </p>
   */
  xlWK3FM3(32),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  xlWKS(4),
  /**
   * <p>
   * The value of this constant is 28
   * </p>
   */
  xlWorks2FarEast(28),
  /**
   * <p>
   * The value of this constant is 34
   * </p>
   */
  xlWQ1(34),
  /**
   * <p>
   * The value of this constant is 40
   * </p>
   */
  xlWJ3(40),
  /**
   * <p>
   * The value of this constant is 41
   * </p>
   */
  xlWJ3FJ3(41),
  /**
   * <p>
   * The value of this constant is 42
   * </p>
   */
  xlUnicodeText(42),
  /**
   * <p>
   * The value of this constant is 44
   * </p>
   */
  xlHtml(44),
  /**
   * <p>
   * The value of this constant is 45
   * </p>
   */
  xlWebArchive(45),
  /**
   * <p>
   * The value of this constant is 46
   * </p>
   */
  xlXMLSpreadsheet(46),
  /**
   * <p>
   * The value of this constant is 50
   * </p>
   */
  xlExcel12(50),
  /**
   * <p>
   * The value of this constant is 51
   * </p>
   */
  xlOpenXMLWorkbook(51),
  /**
   * <p>
   * The value of this constant is 52
   * </p>
   */
  xlOpenXMLWorkbookMacroEnabled(52),
  /**
   * <p>
   * The value of this constant is 53
   * </p>
   */
  xlOpenXMLTemplateMacroEnabled(53),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  xlTemplate8(17),
  /**
   * <p>
   * The value of this constant is 54
   * </p>
   */
  xlOpenXMLTemplate(54),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  xlAddIn8(18),
  /**
   * <p>
   * The value of this constant is 55
   * </p>
   */
  xlOpenXMLAddIn(55),
  /**
   * <p>
   * The value of this constant is 56
   * </p>
   */
  xlExcel8(56),
  /**
   * <p>
   * The value of this constant is 60
   * </p>
   */
  xlOpenDocumentSpreadsheet(60),
  /**
   * <p>
   * The value of this constant is 61
   * </p>
   */
  xlOpenXMLStrictWorkbook(61),
  /**
   * <p>
   * The value of this constant is 62
   * </p>
   */
  xlCSVUTF8(62),
  /**
   * <p>
   * The value of this constant is 51
   * </p>
   */
  xlWorkbookDefault(51),
  ;

  private final int value;
  XlFileFormat(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
