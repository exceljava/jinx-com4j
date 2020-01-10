package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ListObject extends Com4jObject {
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
   */

  @DISPID(117)
  void delete();


  /**
   * @param target Mandatory java.lang.Object parameter.
   * @param linkSource Mandatory boolean parameter.
   */

  @DISPID(1895)
  java.lang.String publish(
    java.lang.Object target,
    boolean linkSource);


  /**
   */

  @DISPID(1417)
  void refresh();


  /**
   */

  @DISPID(2308)
  void unlink();


  /**
   */

  @DISPID(2309)
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

  @DISPID(2310)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {com.exceljava.com4j.excel.XlListConflict.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  void updateChanges();

  /**
   * @param iConflictType Optional parameter. Default value is 0
   */

  @DISPID(2310)
  void updateChanges(
    @Optional @DefaultValue("0") com.exceljava.com4j.excel.XlListConflict iConflictType);


  /**
   * @param range Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(256)
  void resize(
    com.exceljava.com4j.excel.Range range);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  java.lang.String get_Default();


  /**
   * <p>
   * Getter method for the COM property "Active"
   * </p>
   */

  @DISPID(2312)
  @PropGet
  boolean getActive();


  /**
   * <p>
   * Getter method for the COM property "DataBodyRange"
   * </p>
   */

  @DISPID(705)
  @PropGet
  com.exceljava.com4j.excel.Range getDataBodyRange();


  /**
   * <p>
   * Getter method for the COM property "DisplayRightToLeft"
   * </p>
   */

  @DISPID(1774)
  @PropGet
  boolean getDisplayRightToLeft();


  /**
   * <p>
   * Getter method for the COM property "HeaderRowRange"
   * </p>
   */

  @DISPID(2313)
  @PropGet
  com.exceljava.com4j.excel.Range getHeaderRowRange();


  /**
   * <p>
   * Getter method for the COM property "InsertRowRange"
   * </p>
   */

  @DISPID(2314)
  @PropGet
  com.exceljava.com4j.excel.Range getInsertRowRange();


  /**
   * <p>
   * Getter method for the COM property "ListColumns"
   * </p>
   */

  @DISPID(2315)
  @PropGet
  com.exceljava.com4j.excel.ListColumns getListColumns();


  /**
   * <p>
   * Getter method for the COM property "ListRows"
   * </p>
   */

  @DISPID(2316)
  @PropGet
  com.exceljava.com4j.excel.ListRows getListRows();


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
   * Getter method for the COM property "QueryTable"
   * </p>
   */

  @DISPID(1386)
  @PropGet
  com.exceljava.com4j.excel._QueryTable getQueryTable();


  /**
   * <p>
   * Getter method for the COM property "Range"
   * </p>
   */

  @DISPID(197)
  @PropGet
  com.exceljava.com4j.excel.Range getRange();


  /**
   * <p>
   * Getter method for the COM property "ShowAutoFilter"
   * </p>
   */

  @DISPID(2317)
  @PropGet
  boolean getShowAutoFilter();


  /**
   * <p>
   * Setter method for the COM property "ShowAutoFilter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2317)
  @PropPut
  void setShowAutoFilter(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowTotals"
   * </p>
   */

  @DISPID(2318)
  @PropGet
  boolean getShowTotals();


  /**
   * <p>
   * Setter method for the COM property "ShowTotals"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2318)
  @PropPut
  void setShowTotals(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SourceType"
   * </p>
   */

  @DISPID(685)
  @PropGet
  com.exceljava.com4j.excel.XlListObjectSourceType getSourceType();


  /**
   * <p>
   * Getter method for the COM property "TotalsRowRange"
   * </p>
   */

  @DISPID(2319)
  @PropGet
  com.exceljava.com4j.excel.Range getTotalsRowRange();


  /**
   * <p>
   * Getter method for the COM property "SharePointURL"
   * </p>
   */

  @DISPID(2320)
  @PropGet
  java.lang.String getSharePointURL();


  /**
   * <p>
   * Getter method for the COM property "XmlMap"
   * </p>
   */

  @DISPID(2253)
  @PropGet
  com.exceljava.com4j.excel.XmlMap getXmlMap();


  /**
   * <p>
   * Getter method for the COM property "DisplayName"
   * </p>
   */

  @DISPID(2677)
  @PropGet
  java.lang.String getDisplayName();


  /**
   * <p>
   * Setter method for the COM property "DisplayName"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(2677)
  @PropPut
  void setDisplayName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowHeaders"
   * </p>
   */

  @DISPID(2678)
  @PropGet
  boolean getShowHeaders();


  /**
   * <p>
   * Setter method for the COM property "ShowHeaders"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2678)
  @PropPut
  void setShowHeaders(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "_AutoFilter"
   * </p>
   */

  @DISPID(793)
  @PropGet
  com.exceljava.com4j.excel.AutoFilter get_AutoFilter();


  /**
   * <p>
   * Getter method for the COM property "TableStyle"
   * </p>
   */

  @DISPID(1504)
  @PropGet
  java.lang.Object getTableStyle();


  /**
   * <p>
   * Setter method for the COM property "TableStyle"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(1504)
  @PropPut
  void setTableStyle(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowTableStyleFirstColumn"
   * </p>
   */

  @DISPID(2679)
  @PropGet
  boolean getShowTableStyleFirstColumn();


  /**
   * <p>
   * Setter method for the COM property "ShowTableStyleFirstColumn"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2679)
  @PropPut
  void setShowTableStyleFirstColumn(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowTableStyleLastColumn"
   * </p>
   */

  @DISPID(2563)
  @PropGet
  boolean getShowTableStyleLastColumn();


  /**
   * <p>
   * Setter method for the COM property "ShowTableStyleLastColumn"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2563)
  @PropPut
  void setShowTableStyleLastColumn(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowTableStyleRowStripes"
   * </p>
   */

  @DISPID(2564)
  @PropGet
  boolean getShowTableStyleRowStripes();


  /**
   * <p>
   * Setter method for the COM property "ShowTableStyleRowStripes"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2564)
  @PropPut
  void setShowTableStyleRowStripes(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowTableStyleColumnStripes"
   * </p>
   */

  @DISPID(2565)
  @PropGet
  boolean getShowTableStyleColumnStripes();


  /**
   * <p>
   * Setter method for the COM property "ShowTableStyleColumnStripes"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2565)
  @PropPut
  void setShowTableStyleColumnStripes(
    boolean rhs);


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
   * Getter method for the COM property "Comment"
   * </p>
   */

  @DISPID(910)
  @PropGet
  java.lang.String getComment();


  /**
   * <p>
   * Setter method for the COM property "Comment"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(910)
  @PropPut
  void setComment(
    java.lang.String rhs);


  /**
   */

  @DISPID(2680)
  void exportToVisio();


  /**
   * <p>
   * Getter method for the COM property "AlternativeText"
   * </p>
   */

  @DISPID(1891)
  @PropGet
  java.lang.String getAlternativeText();


  /**
   * <p>
   * Setter method for the COM property "AlternativeText"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1891)
  @PropPut
  void setAlternativeText(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Summary"
   * </p>
   */

  @DISPID(273)
  @PropGet
  java.lang.String getSummary();


  /**
   * <p>
   * Setter method for the COM property "Summary"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(273)
  @PropPut
  void setSummary(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "TableObject"
   * </p>
   */

  @DISPID(3095)
  @PropGet
  com.exceljava.com4j.excel.TableObject getTableObject();


  /**
   * <p>
   * Getter method for the COM property "Slicers"
   * </p>
   */

  @DISPID(2881)
  @PropGet
  com.exceljava.com4j.excel.Slicers getSlicers();


  /**
   * <p>
   * Getter method for the COM property "ShowAutoFilterDropDown"
   * </p>
   */

  @DISPID(3096)
  @PropGet
  boolean getShowAutoFilterDropDown();


  /**
   * <p>
   * Setter method for the COM property "ShowAutoFilterDropDown"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(3096)
  @PropPut
  void setShowAutoFilterDropDown(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AutoFilter"
   * </p>
   */

  @DISPID(3289)
  @PropGet
  com.exceljava.com4j.excel.AutoFilter getAutoFilter();


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
