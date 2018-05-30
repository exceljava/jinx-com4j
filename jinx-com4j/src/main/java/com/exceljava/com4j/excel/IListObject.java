package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024471-0001-0000-C000-000000000046}")
public interface IListObject extends Com4jObject {
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
   */

  @VTID(10)
  void delete();


  /**
   * @param target Mandatory java.lang.Object parameter.
   * @param linkSource Mandatory boolean parameter.
   * @return  Returns a value of type java.lang.String
   */

  @VTID(11)
  java.lang.String publish(
    @MarshalAs(NativeType.VARIANT) java.lang.Object target,
    boolean linkSource);


  /**
   */

  @VTID(12)
  void refresh();


  /**
   */

  @VTID(13)
  void unlink();


  /**
   */

  @VTID(14)
  void unlist();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlListConflict parameter iConflictType is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * updateChanges(0);
   * </code>
   * </p>
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {com.exceljava.com4j.excel.XlListConflict.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  void updateChanges();

  /**
   * @param iConflictType Optional parameter. Default value is 0
   */

  @VTID(15)
  void updateChanges(
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlListConflict iConflictType);


  /**
   * @param range Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @VTID(16)
  void resize(
    com.exceljava.com4j.excel.Range range);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(17)
  @DefaultMethod
  java.lang.String get_Default();


  /**
   * <p>
   * Getter method for the COM property "Active"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(18)
  boolean getActive();


  /**
   * <p>
   * Getter method for the COM property "DataBodyRange"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(19)
  com.exceljava.com4j.excel.Range getDataBodyRange();


  /**
   * <p>
   * Getter method for the COM property "DisplayRightToLeft"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(20)
  boolean getDisplayRightToLeft();


  /**
   * <p>
   * Getter method for the COM property "HeaderRowRange"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(21)
  com.exceljava.com4j.excel.Range getHeaderRowRange();


  /**
   * <p>
   * Getter method for the COM property "InsertRowRange"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(22)
  com.exceljava.com4j.excel.Range getInsertRowRange();


  /**
   * <p>
   * Getter method for the COM property "ListColumns"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ListColumns
   */

  @VTID(23)
  com.exceljava.com4j.excel.ListColumns getListColumns();


  /**
   * <p>
   * Getter method for the COM property "ListRows"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ListRows
   */

  @VTID(24)
  com.exceljava.com4j.excel.ListRows getListRows();


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(25)
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(26)
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "QueryTable"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._QueryTable
   */

  @VTID(27)
  com.exceljava.com4j.excel._QueryTable getQueryTable();


  /**
   * <p>
   * Getter method for the COM property "Range"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(28)
  com.exceljava.com4j.excel.Range getRange();


  /**
   * <p>
   * Getter method for the COM property "ShowAutoFilter"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(29)
  boolean getShowAutoFilter();


  /**
   * <p>
   * Setter method for the COM property "ShowAutoFilter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(30)
  void setShowAutoFilter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowTotals"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(31)
  boolean getShowTotals();


  /**
   * <p>
   * Setter method for the COM property "ShowTotals"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(32)
  void setShowTotals(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SourceType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlListObjectSourceType
   */

  @VTID(33)
  com.exceljava.com4j.excel.XlListObjectSourceType getSourceType();


  /**
   * <p>
   * Getter method for the COM property "TotalsRowRange"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(34)
  com.exceljava.com4j.excel.Range getTotalsRowRange();


  /**
   * <p>
   * Getter method for the COM property "SharePointURL"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(35)
  java.lang.String getSharePointURL();


  /**
   * <p>
   * Getter method for the COM property "XmlMap"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XmlMap
   */

  @VTID(36)
  com.exceljava.com4j.excel.XmlMap getXmlMap();


  /**
   * <p>
   * Getter method for the COM property "DisplayName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(37)
  java.lang.String getDisplayName();


  /**
   * <p>
   * Setter method for the COM property "DisplayName"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(38)
  void setDisplayName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowHeaders"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(39)
  boolean getShowHeaders();


  /**
   * <p>
   * Setter method for the COM property "ShowHeaders"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(40)
  void setShowHeaders(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AutoFilter"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.AutoFilter
   */

  @VTID(41)
  com.exceljava.com4j.excel.AutoFilter getAutoFilter();


  /**
   * <p>
   * Getter method for the COM property "TableStyle"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(42)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getTableStyle();


  /**
   * <p>
   * Setter method for the COM property "TableStyle"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(43)
  void setTableStyle(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowTableStyleFirstColumn"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(44)
  boolean getShowTableStyleFirstColumn();


  /**
   * <p>
   * Setter method for the COM property "ShowTableStyleFirstColumn"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(45)
  void setShowTableStyleFirstColumn(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowTableStyleLastColumn"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(46)
  boolean getShowTableStyleLastColumn();


  /**
   * <p>
   * Setter method for the COM property "ShowTableStyleLastColumn"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(47)
  void setShowTableStyleLastColumn(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowTableStyleRowStripes"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(48)
  boolean getShowTableStyleRowStripes();


  /**
   * <p>
   * Setter method for the COM property "ShowTableStyleRowStripes"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(49)
  void setShowTableStyleRowStripes(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowTableStyleColumnStripes"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(50)
  boolean getShowTableStyleColumnStripes();


  /**
   * <p>
   * Setter method for the COM property "ShowTableStyleColumnStripes"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(51)
  void setShowTableStyleColumnStripes(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Sort"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Sort
   */

  @VTID(52)
  com.exceljava.com4j.excel.Sort getSort();


  /**
   * <p>
   * Getter method for the COM property "Comment"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(53)
  java.lang.String getComment();


  /**
   * <p>
   * Setter method for the COM property "Comment"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(54)
  void setComment(
    java.lang.String rhs);


  /**
   */

  @VTID(55)
  void exportToVisio();


  /**
   * <p>
   * Getter method for the COM property "AlternativeText"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(56)
  java.lang.String getAlternativeText();


  /**
   * <p>
   * Setter method for the COM property "AlternativeText"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(57)
  void setAlternativeText(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Summary"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(58)
  java.lang.String getSummary();


  /**
   * <p>
   * Setter method for the COM property "Summary"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(59)
  void setSummary(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "TableObject"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.TableObject
   */

  @VTID(60)
  com.exceljava.com4j.excel.TableObject getTableObject();


  /**
   * <p>
   * Getter method for the COM property "Slicers"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Slicers
   */

  @VTID(61)
  com.exceljava.com4j.excel.Slicers getSlicers();


  /**
   * <p>
   * Getter method for the COM property "ShowAutoFilterDropDown"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(62)
  boolean getShowAutoFilterDropDown();


  /**
   * <p>
   * Setter method for the COM property "ShowAutoFilterDropDown"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(63)
  void setShowAutoFilterDropDown(
    boolean rhs);


  // Properties:
}
