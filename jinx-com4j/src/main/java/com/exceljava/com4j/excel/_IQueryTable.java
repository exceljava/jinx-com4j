package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024428-0001-0000-C000-000000000046}")
public interface _IQueryTable extends Com4jObject {
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
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(10)
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(11)
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "FieldNames"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(12)
  boolean getFieldNames();


  /**
   * <p>
   * Setter method for the COM property "FieldNames"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(13)
  void setFieldNames(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "RowNumbers"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(14)
  boolean getRowNumbers();


  /**
   * <p>
   * Setter method for the COM property "RowNumbers"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(15)
  void setRowNumbers(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "FillAdjacentFormulas"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(16)
  boolean getFillAdjacentFormulas();


  /**
   * <p>
   * Setter method for the COM property "FillAdjacentFormulas"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(17)
  void setFillAdjacentFormulas(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasAutoFormat"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(18)
  boolean getHasAutoFormat();


  /**
   * <p>
   * Setter method for the COM property "HasAutoFormat"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(19)
  void setHasAutoFormat(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "RefreshOnFileOpen"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(20)
  boolean getRefreshOnFileOpen();


  /**
   * <p>
   * Setter method for the COM property "RefreshOnFileOpen"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(21)
  void setRefreshOnFileOpen(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Refreshing"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(22)
  boolean getRefreshing();


  /**
   * <p>
   * Getter method for the COM property "FetchedRowOverflow"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(23)
  boolean getFetchedRowOverflow();


  /**
   * <p>
   * Getter method for the COM property "BackgroundQuery"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(24)
  boolean getBackgroundQuery();


  /**
   * <p>
   * Setter method for the COM property "BackgroundQuery"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(25)
  void setBackgroundQuery(
    boolean rhs);


  /**
   */

  @VTID(26)
  void cancelRefresh();


  /**
   * <p>
   * Getter method for the COM property "RefreshStyle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCellInsertionMode
   */

  @VTID(27)
  com.exceljava.com4j.excel.XlCellInsertionMode getRefreshStyle();


  /**
   * <p>
   * Setter method for the COM property "RefreshStyle"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlCellInsertionMode parameter.
   */

  @VTID(28)
  void setRefreshStyle(
    com.exceljava.com4j.excel.XlCellInsertionMode rhs);


  /**
   * <p>
   * Getter method for the COM property "EnableRefresh"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(29)
  boolean getEnableRefresh();


  /**
   * <p>
   * Setter method for the COM property "EnableRefresh"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(30)
  void setEnableRefresh(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SavePassword"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(31)
  boolean getSavePassword();


  /**
   * <p>
   * Setter method for the COM property "SavePassword"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(32)
  void setSavePassword(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Destination"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(33)
  com.exceljava.com4j.excel.Range getDestination();


  /**
   * <p>
   * Getter method for the COM property "Connection"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(34)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getConnection();


  /**
   * <p>
   * Setter method for the COM property "Connection"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(35)
  void setConnection(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Sql"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(36)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getSql();


  /**
   * <p>
   * Setter method for the COM property "Sql"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(37)
  void setSql(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "PostText"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(38)
  java.lang.String getPostText();


  /**
   * <p>
   * Setter method for the COM property "PostText"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(39)
  void setPostText(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "ResultRange"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(40)
  com.exceljava.com4j.excel.Range getResultRange();


  /**
   */

  @VTID(41)
  void delete();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * refresh(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(42)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  boolean refresh();

  /**
   * @param backgroundQuery Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @VTID(42)
  boolean refresh(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object backgroundQuery);


  /**
   * <p>
   * Getter method for the COM property "Parameters"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Parameters
   */

  @VTID(43)
  com.exceljava.com4j.excel.Parameters getParameters();


  /**
   * <p>
   * Getter method for the COM property "Recordset"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(44)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getRecordset();


  /**
   * <p>
   * Setter method for the COM property "Recordset"
   * </p>
   * @param rhs Mandatory com4j.Com4jObject parameter.
   */

  @VTID(45)
  void setRecordset(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject rhs);


  /**
   * <p>
   * Getter method for the COM property "SaveData"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(46)
  boolean getSaveData();


  /**
   * <p>
   * Setter method for the COM property "SaveData"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(47)
  void setSaveData(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TablesOnlyFromHTML"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(48)
  boolean getTablesOnlyFromHTML();


  /**
   * <p>
   * Setter method for the COM property "TablesOnlyFromHTML"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(49)
  void setTablesOnlyFromHTML(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "EnableEditing"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(50)
  boolean getEnableEditing();


  /**
   * <p>
   * Setter method for the COM property "EnableEditing"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(51)
  void setEnableEditing(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFilePlatform"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(52)
  int getTextFilePlatform();


  /**
   * <p>
   * Setter method for the COM property "TextFilePlatform"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(53)
  void setTextFilePlatform(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileStartRow"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(54)
  int getTextFileStartRow();


  /**
   * <p>
   * Setter method for the COM property "TextFileStartRow"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(55)
  void setTextFileStartRow(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileParseType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlTextParsingType
   */

  @VTID(56)
  com.exceljava.com4j.excel.XlTextParsingType getTextFileParseType();


  /**
   * <p>
   * Setter method for the COM property "TextFileParseType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlTextParsingType parameter.
   */

  @VTID(57)
  void setTextFileParseType(
    com.exceljava.com4j.excel.XlTextParsingType rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileTextQualifier"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlTextQualifier
   */

  @VTID(58)
  com.exceljava.com4j.excel.XlTextQualifier getTextFileTextQualifier();


  /**
   * <p>
   * Setter method for the COM property "TextFileTextQualifier"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlTextQualifier parameter.
   */

  @VTID(59)
  void setTextFileTextQualifier(
    com.exceljava.com4j.excel.XlTextQualifier rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileConsecutiveDelimiter"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(60)
  boolean getTextFileConsecutiveDelimiter();


  /**
   * <p>
   * Setter method for the COM property "TextFileConsecutiveDelimiter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(61)
  void setTextFileConsecutiveDelimiter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileTabDelimiter"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(62)
  boolean getTextFileTabDelimiter();


  /**
   * <p>
   * Setter method for the COM property "TextFileTabDelimiter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(63)
  void setTextFileTabDelimiter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileSemicolonDelimiter"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(64)
  boolean getTextFileSemicolonDelimiter();


  /**
   * <p>
   * Setter method for the COM property "TextFileSemicolonDelimiter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(65)
  void setTextFileSemicolonDelimiter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileCommaDelimiter"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(66)
  boolean getTextFileCommaDelimiter();


  /**
   * <p>
   * Setter method for the COM property "TextFileCommaDelimiter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(67)
  void setTextFileCommaDelimiter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileSpaceDelimiter"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(68)
  boolean getTextFileSpaceDelimiter();


  /**
   * <p>
   * Setter method for the COM property "TextFileSpaceDelimiter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(69)
  void setTextFileSpaceDelimiter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileOtherDelimiter"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(70)
  java.lang.String getTextFileOtherDelimiter();


  /**
   * <p>
   * Setter method for the COM property "TextFileOtherDelimiter"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(71)
  void setTextFileOtherDelimiter(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileColumnDataTypes"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(72)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getTextFileColumnDataTypes();


  /**
   * <p>
   * Setter method for the COM property "TextFileColumnDataTypes"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(73)
  void setTextFileColumnDataTypes(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileFixedColumnWidths"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(74)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getTextFileFixedColumnWidths();


  /**
   * <p>
   * Setter method for the COM property "TextFileFixedColumnWidths"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(75)
  void setTextFileFixedColumnWidths(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "PreserveColumnInfo"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(76)
  boolean getPreserveColumnInfo();


  /**
   * <p>
   * Setter method for the COM property "PreserveColumnInfo"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(77)
  void setPreserveColumnInfo(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "PreserveFormatting"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(78)
  boolean getPreserveFormatting();


  /**
   * <p>
   * Setter method for the COM property "PreserveFormatting"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(79)
  void setPreserveFormatting(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AdjustColumnWidth"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(80)
  boolean getAdjustColumnWidth();


  /**
   * <p>
   * Setter method for the COM property "AdjustColumnWidth"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(81)
  void setAdjustColumnWidth(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "CommandText"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(82)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getCommandText();


  /**
   * <p>
   * Setter method for the COM property "CommandText"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(83)
  void setCommandText(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "CommandType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCmdType
   */

  @VTID(84)
  com.exceljava.com4j.excel.XlCmdType getCommandType();


  /**
   * <p>
   * Setter method for the COM property "CommandType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlCmdType parameter.
   */

  @VTID(85)
  void setCommandType(
    com.exceljava.com4j.excel.XlCmdType rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFilePromptOnRefresh"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(86)
  boolean getTextFilePromptOnRefresh();


  /**
   * <p>
   * Setter method for the COM property "TextFilePromptOnRefresh"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(87)
  void setTextFilePromptOnRefresh(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "QueryType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlQueryType
   */

  @VTID(88)
  com.exceljava.com4j.excel.XlQueryType getQueryType();


  /**
   * <p>
   * Getter method for the COM property "MaintainConnection"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(89)
  boolean getMaintainConnection();


  /**
   * <p>
   * Setter method for the COM property "MaintainConnection"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(90)
  void setMaintainConnection(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileDecimalSeparator"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(91)
  java.lang.String getTextFileDecimalSeparator();


  /**
   * <p>
   * Setter method for the COM property "TextFileDecimalSeparator"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(92)
  void setTextFileDecimalSeparator(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileThousandsSeparator"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(93)
  java.lang.String getTextFileThousandsSeparator();


  /**
   * <p>
   * Setter method for the COM property "TextFileThousandsSeparator"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(94)
  void setTextFileThousandsSeparator(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "RefreshPeriod"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(95)
  int getRefreshPeriod();


  /**
   * <p>
   * Setter method for the COM property "RefreshPeriod"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(96)
  void setRefreshPeriod(
    int rhs);


  /**
   */

  @VTID(97)
  void resetTimer();


  /**
   * <p>
   * Getter method for the COM property "WebSelectionType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlWebSelectionType
   */

  @VTID(98)
  com.exceljava.com4j.excel.XlWebSelectionType getWebSelectionType();


  /**
   * <p>
   * Setter method for the COM property "WebSelectionType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlWebSelectionType parameter.
   */

  @VTID(99)
  void setWebSelectionType(
    com.exceljava.com4j.excel.XlWebSelectionType rhs);


  /**
   * <p>
   * Getter method for the COM property "WebFormatting"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlWebFormatting
   */

  @VTID(100)
  com.exceljava.com4j.excel.XlWebFormatting getWebFormatting();


  /**
   * <p>
   * Setter method for the COM property "WebFormatting"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlWebFormatting parameter.
   */

  @VTID(101)
  void setWebFormatting(
    com.exceljava.com4j.excel.XlWebFormatting rhs);


  /**
   * <p>
   * Getter method for the COM property "WebTables"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(102)
  java.lang.String getWebTables();


  /**
   * <p>
   * Setter method for the COM property "WebTables"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(103)
  void setWebTables(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "WebPreFormattedTextToColumns"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(104)
  boolean getWebPreFormattedTextToColumns();


  /**
   * <p>
   * Setter method for the COM property "WebPreFormattedTextToColumns"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(105)
  void setWebPreFormattedTextToColumns(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "WebSingleBlockTextImport"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(106)
  boolean getWebSingleBlockTextImport();


  /**
   * <p>
   * Setter method for the COM property "WebSingleBlockTextImport"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(107)
  void setWebSingleBlockTextImport(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "WebDisableDateRecognition"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(108)
  boolean getWebDisableDateRecognition();


  /**
   * <p>
   * Setter method for the COM property "WebDisableDateRecognition"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(109)
  void setWebDisableDateRecognition(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "WebConsecutiveDelimitersAsOne"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(110)
  boolean getWebConsecutiveDelimitersAsOne();


  /**
   * <p>
   * Setter method for the COM property "WebConsecutiveDelimitersAsOne"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(111)
  void setWebConsecutiveDelimitersAsOne(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "WebDisableRedirections"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(112)
  boolean getWebDisableRedirections();


  /**
   * <p>
   * Setter method for the COM property "WebDisableRedirections"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(113)
  void setWebDisableRedirections(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "EditWebPage"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(114)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getEditWebPage();


  /**
   * <p>
   * Setter method for the COM property "EditWebPage"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(115)
  void setEditWebPage(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "SourceConnectionFile"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(116)
  java.lang.String getSourceConnectionFile();


  /**
   * <p>
   * Setter method for the COM property "SourceConnectionFile"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(117)
  void setSourceConnectionFile(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "SourceDataFile"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(118)
  java.lang.String getSourceDataFile();


  /**
   * <p>
   * Setter method for the COM property "SourceDataFile"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(119)
  void setSourceDataFile(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "RobustConnect"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlRobustConnect
   */

  @VTID(120)
  com.exceljava.com4j.excel.XlRobustConnect getRobustConnect();


  /**
   * <p>
   * Setter method for the COM property "RobustConnect"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlRobustConnect parameter.
   */

  @VTID(121)
  void setRobustConnect(
    com.exceljava.com4j.excel.XlRobustConnect rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFileTrailingMinusNumbers"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(122)
  boolean getTextFileTrailingMinusNumbers();


  /**
   * <p>
   * Setter method for the COM property "TextFileTrailingMinusNumbers"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(123)
  void setTextFileTrailingMinusNumbers(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter keywords is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAsODC(odcFileName, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param odcFileName Mandatory java.lang.String parameter.
   */

  @VTID(124)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void saveAsODC(
    java.lang.String odcFileName);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter keywords is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAsODC(odcFileName, description, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param odcFileName Mandatory java.lang.String parameter.
   * @param description Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(124)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void saveAsODC(
    java.lang.String odcFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object description);

  /**
   * @param odcFileName Mandatory java.lang.String parameter.
   * @param description Optional parameter. Default value is com4j.Variant.getMissing()
   * @param keywords Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(124)
  void saveAsODC(
    java.lang.String odcFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object description,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object keywords);


  /**
   * <p>
   * Getter method for the COM property "ListObject"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(125)
  com.exceljava.com4j.excel.ListObject getListObject();


  /**
   * <p>
   * Getter method for the COM property "TextFileVisualLayout"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlTextVisualLayoutType
   */

  @VTID(126)
  com.exceljava.com4j.excel.XlTextVisualLayoutType getTextFileVisualLayout();


  /**
   * <p>
   * Setter method for the COM property "TextFileVisualLayout"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlTextVisualLayoutType parameter.
   */

  @VTID(127)
  void setTextFileVisualLayout(
    com.exceljava.com4j.excel.XlTextVisualLayoutType rhs);


  /**
   * <p>
   * Getter method for the COM property "WorkbookConnection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(128)
  com.exceljava.com4j.excel.WorkbookConnection getWorkbookConnection();


  /**
   * <p>
   * Getter method for the COM property "_Sort"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Sort
   */

  @VTID(129)
  com.exceljava.com4j.excel.Sort get_Sort();


  /**
   * <p>
   * Getter method for the COM property "Sort"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Sort
   */

  @VTID(130)
  com.exceljava.com4j.excel.Sort getSort();


  // Properties:
}
