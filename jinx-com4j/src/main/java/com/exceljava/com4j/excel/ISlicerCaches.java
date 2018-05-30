package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244C3-0001-0000-C000-000000000046}")
public interface ISlicerCaches extends Com4jObject,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.SlicerCache
   */

  @VTID(11)
  com.exceljava.com4j.excel.SlicerCache getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.SlicerCache
   */

  @VTID(12)
  @DefaultMethod
  com.exceljava.com4j.excel.SlicerCache get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(13)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(source, sourceField, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param source Mandatory java.lang.Object parameter.
   * @param sourceField Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.SlicerCache
   */

  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.excel.SlicerCache add(
    @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @MarshalAs(NativeType.VARIANT) java.lang.Object sourceField);

  /**
   * @param source Mandatory java.lang.Object parameter.
   * @param sourceField Mandatory java.lang.Object parameter.
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.SlicerCache
   */

  @VTID(14)
  com.exceljava.com4j.excel.SlicerCache add(
    @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @MarshalAs(NativeType.VARIANT) java.lang.Object sourceField,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter slicerCacheType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add2(source, sourceField, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param source Mandatory java.lang.Object parameter.
   * @param sourceField Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.SlicerCache
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=4)
  com.exceljava.com4j.excel.SlicerCache add2(
    @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @MarshalAs(NativeType.VARIANT) java.lang.Object sourceField);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter slicerCacheType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add2(source, sourceField, name, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param source Mandatory java.lang.Object parameter.
   * @param sourceField Mandatory java.lang.Object parameter.
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.SlicerCache
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=4)
  com.exceljava.com4j.excel.SlicerCache add2(
    @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @MarshalAs(NativeType.VARIANT) java.lang.Object sourceField,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name);

  /**
   * @param source Mandatory java.lang.Object parameter.
   * @param sourceField Mandatory java.lang.Object parameter.
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @param slicerCacheType Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.SlicerCache
   */

  @VTID(15)
  com.exceljava.com4j.excel.SlicerCache add2(
    @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @MarshalAs(NativeType.VARIANT) java.lang.Object sourceField,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object slicerCacheType);


  // Properties:
}
