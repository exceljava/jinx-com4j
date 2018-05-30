package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020876-0001-0000-C000-000000000046}")
public interface IPivotItem extends Com4jObject {
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
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotField
   */

  @VTID(9)
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
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getChildItems();

  /**
   * <p>
   * Getter method for the COM property "ChildItems"
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(10)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getChildItems(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "DataRange"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(11)
  com.exceljava.com4j.excel.Range getDataRange();


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(12)
  @DefaultMethod
  java.lang.String get_Default();


  /**
   * <p>
   * Setter method for the COM property "_Default"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(13)
  @DefaultMethod
  void set_Default(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "LabelRange"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(14)
  com.exceljava.com4j.excel.Range getLabelRange();


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(15)
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(16)
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "ParentItem"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotItem
   */

  @VTID(17)
  com.exceljava.com4j.excel.PivotItem getParentItem();


  /**
   * <p>
   * Getter method for the COM property "ParentShowDetail"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(18)
  boolean getParentShowDetail();


  /**
   * <p>
   * Getter method for the COM property "Position"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(19)
  int getPosition();


  /**
   * <p>
   * Setter method for the COM property "Position"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(20)
  void setPosition(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowDetail"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(21)
  boolean getShowDetail();


  /**
   * <p>
   * Setter method for the COM property "ShowDetail"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(22)
  void setShowDetail(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SourceName"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(23)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getSourceName();


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(24)
  java.lang.String getValue();


  /**
   * <p>
   * Setter method for the COM property "Value"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(25)
  void setValue(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(26)
  boolean getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(27)
  void setVisible(
    boolean rhs);


  /**
   */

  @VTID(28)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "IsCalculated"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(29)
  boolean getIsCalculated();


  /**
   * <p>
   * Getter method for the COM property "RecordCount"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(30)
  int getRecordCount();


  /**
   * <p>
   * Getter method for the COM property "Formula"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(31)
  java.lang.String getFormula();


  /**
   * <p>
   * Setter method for the COM property "Formula"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(32)
  void setFormula(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Caption"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(33)
  java.lang.String getCaption();


  /**
   * <p>
   * Setter method for the COM property "Caption"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(34)
  void setCaption(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "DrilledDown"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(35)
  boolean getDrilledDown();


  /**
   * <p>
   * Setter method for the COM property "DrilledDown"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(36)
  void setDrilledDown(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "StandardFormula"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(37)
  java.lang.String getStandardFormula();


  /**
   * <p>
   * Setter method for the COM property "StandardFormula"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(38)
  void setStandardFormula(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "SourceNameStandard"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(39)
  java.lang.String getSourceNameStandard();


  /**
   * @param field Mandatory java.lang.String parameter.
   */

  @VTID(40)
  void drillTo(
    java.lang.String field);


  // Properties:
}
