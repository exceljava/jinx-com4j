package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244C1-0001-0000-C000-000000000046}")
public interface IPivotTableChangeList extends Com4jObject,Iterable<Com4jObject> {
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
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ValueChange
   */

  @VTID(10)
  @DefaultMethod
  com.exceljava.com4j.excel.ValueChange get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(11)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ValueChange
   */

  @VTID(12)
  com.exceljava.com4j.excel.ValueChange getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(13)
  int getCount();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter allocationValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allocationMethod is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allocationWeightExpression is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(tuple, value, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param tuple Mandatory java.lang.String parameter.
   * @param value Mandatory double parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ValueChange
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.ValueChange add(
    java.lang.String tuple,
    double value);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter allocationMethod is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allocationWeightExpression is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(tuple, value, allocationValue, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param tuple Mandatory java.lang.String parameter.
   * @param value Mandatory double parameter.
   * @param allocationValue Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ValueChange
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.ValueChange add(
    java.lang.String tuple,
    double value,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allocationValue);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter allocationWeightExpression is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(tuple, value, allocationValue, allocationMethod, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param tuple Mandatory java.lang.String parameter.
   * @param value Mandatory double parameter.
   * @param allocationValue Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allocationMethod Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ValueChange
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.ValueChange add(
    java.lang.String tuple,
    double value,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allocationValue,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allocationMethod);

  /**
   * @param tuple Mandatory java.lang.String parameter.
   * @param value Mandatory double parameter.
   * @param allocationValue Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allocationMethod Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allocationWeightExpression Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ValueChange
   */

  @VTID(14)
  com.exceljava.com4j.excel.ValueChange add(
    java.lang.String tuple,
    double value,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allocationValue,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allocationMethod,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allocationWeightExpression);


  // Properties:
}
