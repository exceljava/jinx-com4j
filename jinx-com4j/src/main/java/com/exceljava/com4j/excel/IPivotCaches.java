package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002441D-0001-0000-C000-000000000046}")
public interface IPivotCaches extends Com4jObject,Iterable<Com4jObject> {
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
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotCache
   */

  @VTID(11)
  com.exceljava.com4j.excel.PivotCache item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotCache
   */

  @VTID(12)
  @DefaultMethod
  com.exceljava.com4j.excel.PivotCache get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   */

  @VTID(13)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter sourceData is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(sourceType, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Mandatory com.exceljava.com4j.excel.XlPivotTableSourceType parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotCache
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.PivotCache add(
    com.exceljava.com4j.excel.XlPivotTableSourceType sourceType);

  /**
   * @param sourceType Mandatory com.exceljava.com4j.excel.XlPivotTableSourceType parameter.
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotCache
   */

  @VTID(14)
  com.exceljava.com4j.excel.PivotCache add(
    com.exceljava.com4j.excel.XlPivotTableSourceType sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter sourceData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter version is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * create(sourceType, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Mandatory com.exceljava.com4j.excel.XlPivotTableSourceType parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotCache
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.excel.PivotCache create(
    com.exceljava.com4j.excel.XlPivotTableSourceType sourceType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter version is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * create(sourceType, sourceData, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param sourceType Mandatory com.exceljava.com4j.excel.XlPivotTableSourceType parameter.
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotCache
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.excel.PivotCache create(
    com.exceljava.com4j.excel.XlPivotTableSourceType sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData);

  /**
   * @param sourceType Mandatory com.exceljava.com4j.excel.XlPivotTableSourceType parameter.
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param version Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotCache
   */

  @VTID(15)
  com.exceljava.com4j.excel.PivotCache create(
    com.exceljava.com4j.excel.XlPivotTableSourceType sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object version);


  // Properties:
}
