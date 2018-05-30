package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002441F-0001-0000-C000-000000000046}")
public interface IPivotFormulas extends Com4jObject,Iterable<Com4jObject> {
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
   * @param formula Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFormula
   */

  @VTID(11)
  com.exceljava.com4j.excel.PivotFormula _Add(
    java.lang.String formula);


  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFormula
   */

  @VTID(12)
  com.exceljava.com4j.excel.PivotFormula item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFormula
   */

  @VTID(13)
  @DefaultMethod
  com.exceljava.com4j.excel.PivotFormula get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   */

  @VTID(14)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter useStandardFormula is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(formula, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param formula Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFormula
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.PivotFormula add(
    java.lang.String formula);

  /**
   * @param formula Mandatory java.lang.String parameter.
   * @param useStandardFormula Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotFormula
   */

  @VTID(15)
  com.exceljava.com4j.excel.PivotFormula add(
    java.lang.String formula,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object useStandardFormula);


  // Properties:
}
