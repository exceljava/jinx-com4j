package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244CD-0001-0000-C000-000000000046}")
public interface IProtectedViewWindow extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(7)
  @DefaultMethod
  java.lang.String get_Default();


  /**
   * <p>
   * Getter method for the COM property "Caption"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(8)
  java.lang.String getCaption();


  /**
   * <p>
   * Setter method for the COM property "Caption"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(9)
  void setCaption(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "EnableResize"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  boolean getEnableResize();


  /**
   * <p>
   * Setter method for the COM property "EnableResize"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(11)
  void setEnableResize(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(12)
  double getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(13)
  void setHeight(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Left"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(14)
  double getLeft();


  /**
   * <p>
   * Setter method for the COM property "Left"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(15)
  void setLeft(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(16)
  double getTop();


  /**
   * <p>
   * Setter method for the COM property "Top"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(17)
  void setTop(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(18)
  double getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(19)
  void setWidth(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(20)
  boolean getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(21)
  void setVisible(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SourceName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(22)
  java.lang.String getSourceName();


  /**
   * <p>
   * Getter method for the COM property "SourcePath"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(23)
  java.lang.String getSourcePath();


  /**
   * <p>
   * Getter method for the COM property "WindowState"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlProtectedViewWindowState
   */

  @VTID(24)
  com.exceljava.com4j.excel.XlProtectedViewWindowState getWindowState();


  /**
   * <p>
   * Setter method for the COM property "WindowState"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlProtectedViewWindowState parameter.
   */

  @VTID(25)
  void setWindowState(
    com.exceljava.com4j.excel.XlProtectedViewWindowState rhs);


  /**
   * <p>
   * Getter method for the COM property "Workbook"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @VTID(26)
  com.exceljava.com4j.excel._Workbook getWorkbook();


  /**
   */

  @VTID(27)
  void activate();


  /**
   * @return  Returns a value of type boolean
   */

  @VTID(28)
  boolean close();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter updateLinks is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * edit(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @VTID(29)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel._Workbook edit();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter updateLinks is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * edit(writeResPassword, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @VTID(29)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel._Workbook edit(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword);

  /**
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param updateLinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Workbook
   */

  @VTID(29)
  com.exceljava.com4j.excel._Workbook edit(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object updateLinks);


  // Properties:
}
