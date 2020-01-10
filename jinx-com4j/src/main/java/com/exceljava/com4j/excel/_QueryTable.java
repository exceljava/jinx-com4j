package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface _QueryTable extends Com4jObject {
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
   * Getter method for the COM property "Name"
   * </p>
   */

  @DISPID(110)
  @PropGet
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(110)
  @PropPut
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "FieldNames"
   * </p>
   */

  @DISPID(1584)
  @PropGet
  boolean getFieldNames();


  /**
   * <p>
   * Setter method for the COM property "FieldNames"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1584)
  @PropPut
  void setFieldNames(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "RowNumbers"
   * </p>
   */

  @DISPID(1585)
  @PropGet
  boolean getRowNumbers();


  /**
   * <p>
   * Setter method for the COM property "RowNumbers"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1585)
  @PropPut
  void setRowNumbers(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "FillAdjacentFormulas"
   * </p>
   */

  @DISPID(1586)
  @PropGet
  boolean getFillAdjacentFormulas();


  /**
   * <p>
   * Setter method for the COM property "FillAdjacentFormulas"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1586)
  @PropPut
  void setFillAdjacentFormulas(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasAutoFormat"
   * </p>
   */

  @DISPID(695)
  @PropGet
  boolean getHasAutoFormat();


  /**
   * <p>
   * Setter method for the COM property "HasAutoFormat"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(695)
  @PropPut
  void setHasAutoFormat(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "RefreshOnFileOpen"
   * </p>
   */

  @DISPID(1479)
  @PropGet
  boolean getRefreshOnFileOpen();


  /**
   * <p>
   * Setter method for the COM property "RefreshOnFileOpen"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1479)
  @PropPut
  void setRefreshOnFileOpen(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Refreshing"
   * </p>
   */

  @DISPID(1587)
  @PropGet
  boolean getRefreshing();


  /**
   * <p>
   * Getter method for the COM property "FetchedRowOverflow"
   * </p>
   */

  @DISPID(1588)
  @PropGet
  boolean getFetchedRowOverflow();


  /**
   * <p>
   * Getter method for the COM property "BackgroundQuery"
   * </p>
   */

  @DISPID(1427)
  @PropGet
  boolean getBackgroundQuery();


  /**
   * <p>
   * Setter method for the COM property "BackgroundQuery"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1427)
  @PropPut
  void setBackgroundQuery(
    boolean rhs);


  /**
   */

  @DISPID(1589)
  void cancelRefresh();


  /**
   * <p>
   * Getter method for the COM property "RefreshStyle"
   * </p>
   */

  @DISPID(1590)
  @PropGet
  com.exceljava.com4j.excel.XlCellInsertionMode getRefreshStyle();


  /**
   * <p>
   * Setter method for the COM property "RefreshStyle"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlCellInsertionMode parameter.
   */

  @DISPID(1590)
  @PropPut
  void setRefreshStyle(
    com.exceljava.com4j.excel.XlCellInsertionMode rhs);


  /**
   * <p>
   * Getter method for the COM property "EnableRefresh"
   * </p>
   */

  @DISPID(1477)
  @PropGet
  boolean getEnableRefresh();


  /**
   * <p>
   * Setter method for the COM property "EnableRefresh"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1477)
  @PropPut
  void setEnableRefresh(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SavePassword"
   * </p>
   */

  @DISPID(1481)
  @PropGet
  boolean getSavePassword();


  /**
   * <p>
   * Setter method for the COM property "SavePassword"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1481)
  @PropPut
  void setSavePassword(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Destination"
   * </p>
   */

  @DISPID(681)
  @PropGet
  com.exceljava.com4j.excel.Range getDestination();


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
   * Getter method for the COM property "Sql"
   * </p>
   */

  @DISPID(1480)
  @PropGet
  java.lang.Object getSql();


  /**
   * <p>
   * Setter method for the COM property "Sql"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1480)
  @PropPut
  void setSql(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "PostText"
   * </p>
   */

  @DISPID(1591)
  @PropGet
  java.lang.String getPostText();


  /**
   * <p>
   * Setter method for the COM property "PostText"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1591)
  @PropPut
  void setPostText(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "ResultRange"
   * </p>
   */

  @DISPID(1592)
  @PropGet
  com.exceljava.com4j.excel.Range getResultRange();


  /**
   */

  @DISPID(117)
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
   */

  @DISPID(1417)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  boolean refresh();

  /**
   * @param backgroundQuery Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1417)
  boolean refresh(
    @Optional java.lang.Object backgroundQuery);


  /**
   * <p>
   * Getter method for the COM property "Parameters"
   * </p>
   */

  @DISPID(1593)
  @PropGet
  com.exceljava.com4j.excel.Parameters getParameters();


  /**
   * <p>
   * Getter method for the COM property "Recordset"
   * </p>
   */

  @DISPID(1165)
  @PropGet
  com4j.Com4jObject getRecordset();


  /**
   * <p>
   * Setter method for the COM property "Recordset"
   * </p>
   * @param rhs Mandatory com4j.Com4jObject parameter.
   */

  @DISPID(1165)
  @PropPut
  void setRecordset(
    com4j.Com4jObject rhs);


  /**
   * <p>
   * Getter method for the COM property "SaveData"
   * </p>
   */

  @DISPID(692)
  @PropGet
  boolean getSaveData();


  /**
   * <p>
   * Setter method for the COM property "SaveData"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(692)
  @PropPut
  void setSaveData(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TablesOnlyFromHTML"
   * </p>
   */

  @DISPID(1594)
  @PropGet
  boolean getTablesOnlyFromHTML();


  /**
   * <p>
   * Setter method for the COM property "TablesOnlyFromHTML"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1594)
  @PropPut
  void setTablesOnlyFromHTML(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "EnableEditing"
   * </p>
   */

  @DISPID(1595)
  @PropGet
  boolean getEnableEditing();


  /**
   * <p>
   * Setter method for the COM property "EnableEditing"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1595)
  @PropPut
  void setEnableEditing(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFilePlatform"
   * </p>
   */

  @DISPID(1855)
  @PropGet
  int getTextFilePlatform();


  /**
   * <p>
   * Setter method for the COM property "TextFilePlatform"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(1855)
  @PropPut
  void setTextFilePlatform(
    int rhs);


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
   * Getter method for the COM property "PreserveColumnInfo"
   * </p>
   */

  @DISPID(1867)
  @PropGet
  boolean getPreserveColumnInfo();


  /**
   * <p>
   * Setter method for the COM property "PreserveColumnInfo"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1867)
  @PropPut
  void setPreserveColumnInfo(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "PreserveFormatting"
   * </p>
   */

  @DISPID(1500)
  @PropGet
  boolean getPreserveFormatting();


  /**
   * <p>
   * Setter method for the COM property "PreserveFormatting"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1500)
  @PropPut
  void setPreserveFormatting(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AdjustColumnWidth"
   * </p>
   */

  @DISPID(1868)
  @PropGet
  boolean getAdjustColumnWidth();


  /**
   * <p>
   * Setter method for the COM property "AdjustColumnWidth"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1868)
  @PropPut
  void setAdjustColumnWidth(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "CommandText"
   * </p>
   */

  @DISPID(1829)
  @PropGet
  java.lang.Object getCommandText();


  /**
   * <p>
   * Setter method for the COM property "CommandText"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1829)
  @PropPut
  void setCommandText(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "CommandType"
   * </p>
   */

  @DISPID(1830)
  @PropGet
  com.exceljava.com4j.excel.XlCmdType getCommandType();


  /**
   * <p>
   * Setter method for the COM property "CommandType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlCmdType parameter.
   */

  @DISPID(1830)
  @PropPut
  void setCommandType(
    com.exceljava.com4j.excel.XlCmdType rhs);


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
   * Getter method for the COM property "QueryType"
   * </p>
   */

  @DISPID(1831)
  @PropGet
  com.exceljava.com4j.excel.XlQueryType getQueryType();


  /**
   * <p>
   * Getter method for the COM property "MaintainConnection"
   * </p>
   */

  @DISPID(1832)
  @PropGet
  boolean getMaintainConnection();


  /**
   * <p>
   * Setter method for the COM property "MaintainConnection"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1832)
  @PropPut
  void setMaintainConnection(
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
   * Getter method for the COM property "RefreshPeriod"
   * </p>
   */

  @DISPID(1833)
  @PropGet
  int getRefreshPeriod();


  /**
   * <p>
   * Setter method for the COM property "RefreshPeriod"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(1833)
  @PropPut
  void setRefreshPeriod(
    int rhs);


  /**
   */

  @DISPID(1834)
  void resetTimer();


  /**
   * <p>
   * Getter method for the COM property "WebSelectionType"
   * </p>
   */

  @DISPID(1872)
  @PropGet
  com.exceljava.com4j.excel.XlWebSelectionType getWebSelectionType();


  /**
   * <p>
   * Setter method for the COM property "WebSelectionType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlWebSelectionType parameter.
   */

  @DISPID(1872)
  @PropPut
  void setWebSelectionType(
    com.exceljava.com4j.excel.XlWebSelectionType rhs);


  /**
   * <p>
   * Getter method for the COM property "WebFormatting"
   * </p>
   */

  @DISPID(1873)
  @PropGet
  com.exceljava.com4j.excel.XlWebFormatting getWebFormatting();


  /**
   * <p>
   * Setter method for the COM property "WebFormatting"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlWebFormatting parameter.
   */

  @DISPID(1873)
  @PropPut
  void setWebFormatting(
    com.exceljava.com4j.excel.XlWebFormatting rhs);


  /**
   * <p>
   * Getter method for the COM property "WebTables"
   * </p>
   */

  @DISPID(1874)
  @PropGet
  java.lang.String getWebTables();


  /**
   * <p>
   * Setter method for the COM property "WebTables"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1874)
  @PropPut
  void setWebTables(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "WebPreFormattedTextToColumns"
   * </p>
   */

  @DISPID(1875)
  @PropGet
  boolean getWebPreFormattedTextToColumns();


  /**
   * <p>
   * Setter method for the COM property "WebPreFormattedTextToColumns"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1875)
  @PropPut
  void setWebPreFormattedTextToColumns(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "WebSingleBlockTextImport"
   * </p>
   */

  @DISPID(1876)
  @PropGet
  boolean getWebSingleBlockTextImport();


  /**
   * <p>
   * Setter method for the COM property "WebSingleBlockTextImport"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1876)
  @PropPut
  void setWebSingleBlockTextImport(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "WebDisableDateRecognition"
   * </p>
   */

  @DISPID(1877)
  @PropGet
  boolean getWebDisableDateRecognition();


  /**
   * <p>
   * Setter method for the COM property "WebDisableDateRecognition"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1877)
  @PropPut
  void setWebDisableDateRecognition(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "WebConsecutiveDelimitersAsOne"
   * </p>
   */

  @DISPID(1878)
  @PropGet
  boolean getWebConsecutiveDelimitersAsOne();


  /**
   * <p>
   * Setter method for the COM property "WebConsecutiveDelimitersAsOne"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1878)
  @PropPut
  void setWebConsecutiveDelimitersAsOne(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "WebDisableRedirections"
   * </p>
   */

  @DISPID(2162)
  @PropGet
  boolean getWebDisableRedirections();


  /**
   * <p>
   * Setter method for the COM property "WebDisableRedirections"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2162)
  @PropPut
  void setWebDisableRedirections(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "EditWebPage"
   * </p>
   */

  @DISPID(2163)
  @PropGet
  java.lang.Object getEditWebPage();


  /**
   * <p>
   * Setter method for the COM property "EditWebPage"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(2163)
  @PropPut
  void setEditWebPage(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "SourceConnectionFile"
   * </p>
   */

  @DISPID(2079)
  @PropGet
  java.lang.String getSourceConnectionFile();


  /**
   * <p>
   * Setter method for the COM property "SourceConnectionFile"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(2079)
  @PropPut
  void setSourceConnectionFile(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "SourceDataFile"
   * </p>
   */

  @DISPID(2080)
  @PropGet
  java.lang.String getSourceDataFile();


  /**
   * <p>
   * Setter method for the COM property "SourceDataFile"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(2080)
  @PropPut
  void setSourceDataFile(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "RobustConnect"
   * </p>
   */

  @DISPID(2081)
  @PropGet
  com.exceljava.com4j.excel.XlRobustConnect getRobustConnect();


  /**
   * <p>
   * Setter method for the COM property "RobustConnect"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlRobustConnect parameter.
   */

  @DISPID(2081)
  @PropPut
  void setRobustConnect(
    com.exceljava.com4j.excel.XlRobustConnect rhs);


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

  @DISPID(2082)
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

  @DISPID(2082)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void saveAsODC(
    java.lang.String odcFileName,
    @Optional java.lang.Object description);

  /**
   * @param odcFileName Mandatory java.lang.String parameter.
   * @param description Optional parameter. Default value is com4j.Variant.getMissing()
   * @param keywords Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2082)
  void saveAsODC(
    java.lang.String odcFileName,
    @Optional java.lang.Object description,
    @Optional java.lang.Object keywords);


  /**
   * <p>
   * Getter method for the COM property "ListObject"
   * </p>
   */

  @DISPID(2257)
  @PropGet
  com.exceljava.com4j.excel.ListObject getListObject();


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


  /**
   * <p>
   * Getter method for the COM property "WorkbookConnection"
   * </p>
   */

  @DISPID(2544)
  @PropGet
  com.exceljava.com4j.excel.WorkbookConnection getWorkbookConnection();


  /**
   * <p>
   * Getter method for the COM property "_Sort"
   * </p>
   */

  @DISPID(880)
  @PropGet
  com.exceljava.com4j.excel.Sort get_Sort();


  /**
   * <p>
   * Getter method for the COM property "Sort"
   * </p>
   */

  @DISPID(3288)
  @PropGet
  com.exceljava.com4j.excel.Sort getSort();


  // Properties:
}
