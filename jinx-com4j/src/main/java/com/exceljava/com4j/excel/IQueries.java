package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244EC-0001-0000-C000-000000000046}")
public interface IQueries extends Com4jObject,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(10)
  int getCount();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter description is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(name, formula, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookQuery
   */

  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.excel.WorkbookQuery add(
    java.lang.String name,
    java.lang.String formula);

  /**
   * @param name Mandatory java.lang.String parameter.
   * @param formula Mandatory java.lang.String parameter.
   * @param description Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookQuery
   */

  @VTID(11)
  com.exceljava.com4j.excel.WorkbookQuery add(
    java.lang.String name,
    java.lang.String formula,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object description);


  /**
   * @param nameOrIndex Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookQuery
   */

  @VTID(12)
  com.exceljava.com4j.excel.WorkbookQuery item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object nameOrIndex);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param nameOrIndex Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookQuery
   */

  @VTID(13)
  @DefaultMethod
  com.exceljava.com4j.excel.WorkbookQuery get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object nameOrIndex);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(14)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "FastCombine"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(15)
  boolean getFastCombine();


  /**
   * <p>
   * Setter method for the COM property "FastCombine"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(16)
  void setFastCombine(
    boolean rhs);


  // Properties:
}
