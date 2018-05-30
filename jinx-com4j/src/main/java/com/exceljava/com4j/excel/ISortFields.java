package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244AA-0001-0000-C000-000000000046}")
public interface ISortFields extends Com4jObject,Iterable<Com4jObject> {
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
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter sortOn is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter order is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter customOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dataOption is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(key, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key Mandatory com.exceljava.com4j.excel.Range parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.SortField
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 5}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.SortField add(
    com.exceljava.com4j.excel.Range key);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter order is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter customOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dataOption is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(key, sortOn, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param sortOn Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.SortField
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.SortField add(
    com.exceljava.com4j.excel.Range key,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sortOn);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter customOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter dataOption is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(key, sortOn, order, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param sortOn Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.SortField
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.SortField add(
    com.exceljava.com4j.excel.Range key,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sortOn,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object order);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter dataOption is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(key, sortOn, order, customOrder, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param key Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param sortOn Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is com4j.Variant.getMissing()
   * @param customOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.SortField
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.SortField add(
    com.exceljava.com4j.excel.Range key,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sortOn,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object order,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customOrder);

  /**
   * @param key Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param sortOn Optional parameter. Default value is com4j.Variant.getMissing()
   * @param order Optional parameter. Default value is com4j.Variant.getMissing()
   * @param customOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param dataOption Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.SortField
   */

  @VTID(10)
  com.exceljava.com4j.excel.SortField add(
    com.exceljava.com4j.excel.Range key,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sortOn,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object order,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object dataOption);


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.SortField
   */

  @VTID(11)
  com.exceljava.com4j.excel.SortField getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(12)
  int getCount();


  /**
   */

  @VTID(13)
  void clear();


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.SortField
   */

  @VTID(14)
  @DefaultMethod
  com.exceljava.com4j.excel.SortField get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(15)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
