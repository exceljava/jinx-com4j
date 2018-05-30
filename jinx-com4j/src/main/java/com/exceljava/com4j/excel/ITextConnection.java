package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244D3-0001-0000-C000-000000000046}")
public interface ITextConnection extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Application
   */

  @VTID(7)
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCreator
   */

  @VTID(8)
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Connection"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getConnection();


  /**
   * <p>
   * Setter method for the COM property "Connection"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(11)
  void setConnection(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileHeaderRow"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(12)
  boolean getTextFileHeaderRow();


  /**
   * <p>
   * Setter method for the COM property "TextFileHeaderRow"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(13)
  void setTextFileHeaderRow(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileColumnDataTypes"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(14)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getTextFileColumnDataTypes();


  /**
   * <p>
   * Setter method for the COM property "TextFileColumnDataTypes"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(15)
  void setTextFileColumnDataTypes(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileCommaDelimiter"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(16)
  boolean getTextFileCommaDelimiter();


  /**
   * <p>
   * Setter method for the COM property "TextFileCommaDelimiter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(17)
  void setTextFileCommaDelimiter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileConsecutiveDelimiter"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(18)
  boolean getTextFileConsecutiveDelimiter();


  /**
   * <p>
   * Setter method for the COM property "TextFileConsecutiveDelimiter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(19)
  void setTextFileConsecutiveDelimiter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileDecimalSeparator"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(20)
  java.lang.String getTextFileDecimalSeparator();


  /**
   * <p>
   * Setter method for the COM property "TextFileDecimalSeparator"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(21)
  void setTextFileDecimalSeparator(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileFixedColumnWidths"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(22)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getTextFileFixedColumnWidths();


  /**
   * <p>
   * Setter method for the COM property "TextFileFixedColumnWidths"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(23)
  void setTextFileFixedColumnWidths(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileOtherDelimiter"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(24)
  java.lang.String getTextFileOtherDelimiter();


  /**
   * <p>
   * Setter method for the COM property "TextFileOtherDelimiter"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(25)
  void setTextFileOtherDelimiter(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileParseType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlTextParsingType
   */

  @VTID(26)
  com.exceljava.com4j.excel.XlTextParsingType getTextFileParseType();


  /**
   * <p>
   * Setter method for the COM property "TextFileParseType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlTextParsingType parameter.
   */

  @VTID(27)
  void setTextFileParseType(
    com.exceljava.com4j.excel.XlTextParsingType rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFilePlatform"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlPlatform
   */

  @VTID(28)
  com.exceljava.com4j.excel.XlPlatform getTextFilePlatform();


  /**
   * <p>
   * Setter method for the COM property "TextFilePlatform"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlPlatform parameter.
   */

  @VTID(29)
  void setTextFilePlatform(
    com.exceljava.com4j.excel.XlPlatform rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFilePromptOnRefresh"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(30)
  boolean getTextFilePromptOnRefresh();


  /**
   * <p>
   * Setter method for the COM property "TextFilePromptOnRefresh"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(31)
  void setTextFilePromptOnRefresh(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileSemicolonDelimiter"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(32)
  boolean getTextFileSemicolonDelimiter();


  /**
   * <p>
   * Setter method for the COM property "TextFileSemicolonDelimiter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(33)
  void setTextFileSemicolonDelimiter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileSpaceDelimiter"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(34)
  boolean getTextFileSpaceDelimiter();


  /**
   * <p>
   * Setter method for the COM property "TextFileSpaceDelimiter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(35)
  void setTextFileSpaceDelimiter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileStartRow"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(36)
  int getTextFileStartRow();


  /**
   * <p>
   * Setter method for the COM property "TextFileStartRow"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(37)
  void setTextFileStartRow(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileTabDelimiter"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(38)
  boolean getTextFileTabDelimiter();


  /**
   * <p>
   * Setter method for the COM property "TextFileTabDelimiter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(39)
  void setTextFileTabDelimiter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileTextQualifier"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlTextQualifier
   */

  @VTID(40)
  com.exceljava.com4j.excel.XlTextQualifier getTextFileTextQualifier();


  /**
   * <p>
   * Setter method for the COM property "TextFileTextQualifier"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlTextQualifier parameter.
   */

  @VTID(41)
  void setTextFileTextQualifier(
    com.exceljava.com4j.excel.XlTextQualifier rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileThousandsSeparator"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(42)
  java.lang.String getTextFileThousandsSeparator();


  /**
   * <p>
   * Setter method for the COM property "TextFileThousandsSeparator"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(43)
  void setTextFileThousandsSeparator(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileTrailingMinusNumbers"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(44)
  boolean getTextFileTrailingMinusNumbers();


  /**
   * <p>
   * Setter method for the COM property "TextFileTrailingMinusNumbers"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(45)
  void setTextFileTrailingMinusNumbers(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileVisualLayout"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlTextVisualLayoutType
   */

  @VTID(46)
  com.exceljava.com4j.excel.XlTextVisualLayoutType getTextFileVisualLayout();


  /**
   * <p>
   * Setter method for the COM property "TextFileVisualLayout"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlTextVisualLayoutType parameter.
   */

  @VTID(47)
  void setTextFileVisualLayout(
    com.exceljava.com4j.excel.XlTextVisualLayoutType rhs);


  // Properties:
}
