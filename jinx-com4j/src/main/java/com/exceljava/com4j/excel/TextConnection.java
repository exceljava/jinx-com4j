package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface TextConnection extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   */

  @DISPID(148)
  @PropGet
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   */

  @DISPID(149)
  @PropGet
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   */

  @DISPID(150)
  @PropGet
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Connection"
   * </p>
   */

  @DISPID(1432)
  @PropGet
  java.lang.Object getConnection();


  /**
   * <p>
   * Setter method for the COM property "Connection"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1432)
  @PropPut
  void setConnection(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileHeaderRow"
   * </p>
   */

  @DISPID(3118)
  @PropGet
  boolean getTextFileHeaderRow();


  /**
   * <p>
   * Setter method for the COM property "TextFileHeaderRow"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(3118)
  @PropPut
  void setTextFileHeaderRow(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileColumnDataTypes"
   * </p>
   */

  @DISPID(1865)
  @PropGet
  java.lang.Object getTextFileColumnDataTypes();


  /**
   * <p>
   * Setter method for the COM property "TextFileColumnDataTypes"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1865)
  @PropPut
  void setTextFileColumnDataTypes(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileCommaDelimiter"
   * </p>
   */

  @DISPID(1862)
  @PropGet
  boolean getTextFileCommaDelimiter();


  /**
   * <p>
   * Setter method for the COM property "TextFileCommaDelimiter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1862)
  @PropPut
  void setTextFileCommaDelimiter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileConsecutiveDelimiter"
   * </p>
   */

  @DISPID(1859)
  @PropGet
  boolean getTextFileConsecutiveDelimiter();


  /**
   * <p>
   * Setter method for the COM property "TextFileConsecutiveDelimiter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1859)
  @PropPut
  void setTextFileConsecutiveDelimiter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileDecimalSeparator"
   * </p>
   */

  @DISPID(1870)
  @PropGet
  java.lang.String getTextFileDecimalSeparator();


  /**
   * <p>
   * Setter method for the COM property "TextFileDecimalSeparator"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1870)
  @PropPut
  void setTextFileDecimalSeparator(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileFixedColumnWidths"
   * </p>
   */

  @DISPID(1866)
  @PropGet
  java.lang.Object getTextFileFixedColumnWidths();


  /**
   * <p>
   * Setter method for the COM property "TextFileFixedColumnWidths"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1866)
  @PropPut
  void setTextFileFixedColumnWidths(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileOtherDelimiter"
   * </p>
   */

  @DISPID(1864)
  @PropGet
  java.lang.String getTextFileOtherDelimiter();


  /**
   * <p>
   * Setter method for the COM property "TextFileOtherDelimiter"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1864)
  @PropPut
  void setTextFileOtherDelimiter(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileParseType"
   * </p>
   */

  @DISPID(1857)
  @PropGet
  com.exceljava.com4j.excel.XlTextParsingType getTextFileParseType();


  /**
   * <p>
   * Setter method for the COM property "TextFileParseType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlTextParsingType parameter.
   */

  @DISPID(1857)
  @PropPut
  void setTextFileParseType(
    com.exceljava.com4j.excel.XlTextParsingType rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFilePlatform"
   * </p>
   */

  @DISPID(1855)
  @PropGet
  com.exceljava.com4j.excel.XlPlatform getTextFilePlatform();


  /**
   * <p>
   * Setter method for the COM property "TextFilePlatform"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlPlatform parameter.
   */

  @DISPID(1855)
  @PropPut
  void setTextFilePlatform(
    com.exceljava.com4j.excel.XlPlatform rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFilePromptOnRefresh"
   * </p>
   */

  @DISPID(1869)
  @PropGet
  boolean getTextFilePromptOnRefresh();


  /**
   * <p>
   * Setter method for the COM property "TextFilePromptOnRefresh"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1869)
  @PropPut
  void setTextFilePromptOnRefresh(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileSemicolonDelimiter"
   * </p>
   */

  @DISPID(1861)
  @PropGet
  boolean getTextFileSemicolonDelimiter();


  /**
   * <p>
   * Setter method for the COM property "TextFileSemicolonDelimiter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1861)
  @PropPut
  void setTextFileSemicolonDelimiter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileSpaceDelimiter"
   * </p>
   */

  @DISPID(1863)
  @PropGet
  boolean getTextFileSpaceDelimiter();


  /**
   * <p>
   * Setter method for the COM property "TextFileSpaceDelimiter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1863)
  @PropPut
  void setTextFileSpaceDelimiter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileStartRow"
   * </p>
   */

  @DISPID(1856)
  @PropGet
  int getTextFileStartRow();


  /**
   * <p>
   * Setter method for the COM property "TextFileStartRow"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(1856)
  @PropPut
  void setTextFileStartRow(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileTabDelimiter"
   * </p>
   */

  @DISPID(1860)
  @PropGet
  boolean getTextFileTabDelimiter();


  /**
   * <p>
   * Setter method for the COM property "TextFileTabDelimiter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1860)
  @PropPut
  void setTextFileTabDelimiter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileTextQualifier"
   * </p>
   */

  @DISPID(1858)
  @PropGet
  com.exceljava.com4j.excel.XlTextQualifier getTextFileTextQualifier();


  /**
   * <p>
   * Setter method for the COM property "TextFileTextQualifier"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlTextQualifier parameter.
   */

  @DISPID(1858)
  @PropPut
  void setTextFileTextQualifier(
    com.exceljava.com4j.excel.XlTextQualifier rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileThousandsSeparator"
   * </p>
   */

  @DISPID(1871)
  @PropGet
  java.lang.String getTextFileThousandsSeparator();


  /**
   * <p>
   * Setter method for the COM property "TextFileThousandsSeparator"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1871)
  @PropPut
  void setTextFileThousandsSeparator(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileTrailingMinusNumbers"
   * </p>
   */

  @DISPID(2164)
  @PropGet
  boolean getTextFileTrailingMinusNumbers();


  /**
   * <p>
   * Setter method for the COM property "TextFileTrailingMinusNumbers"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2164)
  @PropPut
  void setTextFileTrailingMinusNumbers(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileVisualLayout"
   * </p>
   */

  @DISPID(2245)
  @PropGet
  com.exceljava.com4j.excel.XlTextVisualLayoutType getTextFileVisualLayout();


  /**
   * <p>
   * Setter method for the COM property "TextFileVisualLayout"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlTextVisualLayoutType parameter.
   */

  @DISPID(2245)
  @PropPut
  void setTextFileVisualLayout(
    com.exceljava.com4j.excel.XlTextVisualLayoutType rhs);


  // Properties:
}
