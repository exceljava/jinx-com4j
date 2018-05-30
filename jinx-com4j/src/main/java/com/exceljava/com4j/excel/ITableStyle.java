package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244A7-0001-0000-C000-000000000046}")
public interface ITableStyle extends Com4jObject {
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
   * Getter method for the COM property "_Default"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(10)
  @DefaultMethod
  java.lang.String get_Default();


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(11)
  java.lang.String getName();


  /**
   * <p>
   * Getter method for the COM property "NameLocal"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(12)
  java.lang.String getNameLocal();


  /**
   * <p>
   * Getter method for the COM property "BuiltIn"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(13)
  boolean getBuiltIn();


  /**
   * <p>
   * Getter method for the COM property "TableStyleElements"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.TableStyleElements
   */

  @VTID(14)
  com.exceljava.com4j.excel.TableStyleElements getTableStyleElements();


  /**
   * <p>
   * Getter method for the COM property "ShowAsAvailableTableStyle"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(15)
  boolean getShowAsAvailableTableStyle();


  /**
   * <p>
   * Setter method for the COM property "ShowAsAvailableTableStyle"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(16)
  void setShowAsAvailableTableStyle(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowAsAvailablePivotTableStyle"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(17)
  boolean getShowAsAvailablePivotTableStyle();


  /**
   * <p>
   * Setter method for the COM property "ShowAsAvailablePivotTableStyle"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(18)
  void setShowAsAvailablePivotTableStyle(
    boolean rhs);


  /**
   */

  @VTID(19)
  void delete();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter newTableStyleName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * duplicate(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.TableStyle
   */

  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.TableStyle duplicate();

  /**
   * @param newTableStyleName Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.TableStyle
   */

  @VTID(20)
  com.exceljava.com4j.excel.TableStyle duplicate(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object newTableStyleName);


  /**
   * <p>
   * Getter method for the COM property "ShowAsAvailableSlicerStyle"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(21)
  boolean getShowAsAvailableSlicerStyle();


  /**
   * <p>
   * Setter method for the COM property "ShowAsAvailableSlicerStyle"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(22)
  void setShowAsAvailableSlicerStyle(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowAsAvailableTimelineStyle"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(23)
  boolean getShowAsAvailableTimelineStyle();


  /**
   * <p>
   * Setter method for the COM property "ShowAsAvailableTimelineStyle"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(24)
  void setShowAsAvailableTimelineStyle(
    boolean rhs);


  // Properties:
}
