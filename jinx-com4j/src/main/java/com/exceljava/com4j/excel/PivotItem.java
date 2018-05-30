package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface PivotItem extends Com4jObject {
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
  com.exceljava.com4j.excel.PivotField getParent();


  /**
   * <p>
   * Getter method for the COM property "ChildItems"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getChildItems(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(730)
  @PropGet
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.Object getChildItems();

  /**
   * <p>
   * Getter method for the COM property "ChildItems"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(730)
  @PropGet
  java.lang.Object getChildItems(
    @Optional java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "DataRange"
   * </p>
   */

  @DISPID(720)
  @PropGet
  com.exceljava.com4j.excel.Range getDataRange();


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
   * Setter method for the COM property "_Default"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(0)
  @PropPut
  @DefaultMethod
  void set_Default(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "LabelRange"
   * </p>
   */

  @DISPID(719)
  @PropGet
  com.exceljava.com4j.excel.Range getLabelRange();


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
   * Getter method for the COM property "ParentItem"
   * </p>
   */

  @DISPID(741)
  @PropGet
  com.exceljava.com4j.excel.PivotItem getParentItem();


  /**
   * <p>
   * Getter method for the COM property "ParentShowDetail"
   * </p>
   */

  @DISPID(739)
  @PropGet
  boolean getParentShowDetail();


  /**
   * <p>
   * Getter method for the COM property "Position"
   * </p>
   */

  @DISPID(133)
  @PropGet
  int getPosition();


  /**
   * <p>
   * Setter method for the COM property "Position"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(133)
  @PropPut
  void setPosition(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowDetail"
   * </p>
   */

  @DISPID(585)
  @PropGet
  boolean getShowDetail();


  /**
   * <p>
   * Setter method for the COM property "ShowDetail"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(585)
  @PropPut
  void setShowDetail(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SourceName"
   * </p>
   */

  @DISPID(721)
  @PropGet
  java.lang.Object getSourceName();


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   */

  @DISPID(6)
  @PropGet
  java.lang.String getValue();


  /**
   * <p>
   * Setter method for the COM property "Value"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(6)
  @PropPut
  void setValue(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   */

  @DISPID(558)
  @PropGet
  boolean getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(558)
  @PropPut
  void setVisible(
    boolean rhs);


  /**
   */

  @DISPID(117)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "IsCalculated"
   * </p>
   */

  @DISPID(1512)
  @PropGet
  boolean getIsCalculated();


  /**
   * <p>
   * Getter method for the COM property "RecordCount"
   * </p>
   */

  @DISPID(1478)
  @PropGet
  int getRecordCount();


  /**
   * <p>
   * Getter method for the COM property "Formula"
   * </p>
   */

  @DISPID(261)
  @PropGet
  java.lang.String getFormula();


  /**
   * <p>
   * Setter method for the COM property "Formula"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(261)
  @PropPut
  void setFormula(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Caption"
   * </p>
   */

  @DISPID(139)
  @PropGet
  java.lang.String getCaption();


  /**
   * <p>
   * Setter method for the COM property "Caption"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(139)
  @PropPut
  void setCaption(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "DrilledDown"
   * </p>
   */

  @DISPID(1850)
  @PropGet
  boolean getDrilledDown();


  /**
   * <p>
   * Setter method for the COM property "DrilledDown"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1850)
  @PropPut
  void setDrilledDown(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "StandardFormula"
   * </p>
   */

  @DISPID(2084)
  @PropGet
  java.lang.String getStandardFormula();


  /**
   * <p>
   * Setter method for the COM property "StandardFormula"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(2084)
  @PropPut
  void setStandardFormula(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "SourceNameStandard"
   * </p>
   */

  @DISPID(2148)
  @PropGet
  java.lang.String getSourceNameStandard();


  /**
   * @param field Mandatory java.lang.String parameter.
   */

  @DISPID(2580)
  void drillTo(
    java.lang.String field);


  // Properties:
}
