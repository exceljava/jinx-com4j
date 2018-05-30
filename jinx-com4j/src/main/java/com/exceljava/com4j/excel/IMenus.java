package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020865-0001-0000-C000-000000000046}")
public interface IMenus extends Com4jObject,Iterable<Com4jObject> {
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
     * <li>java.lang.Object parameter before is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter restore is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(caption, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param caption Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Menu
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.excel.Menu add(
    java.lang.String caption);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter restore is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(caption, before, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param caption Mandatory java.lang.String parameter.
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Menu
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.excel.Menu add(
    java.lang.String caption,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before);

  /**
   * @param caption Mandatory java.lang.String parameter.
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @param restore Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Menu
   */

  @VTID(10)
  com.exceljava.com4j.excel.Menu add(
    java.lang.String caption,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object restore);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(11)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Menu
   */

  @VTID(12)
  @DefaultMethod
  com.exceljava.com4j.excel.Menu get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Menu
   */

  @VTID(13)
  com.exceljava.com4j.excel.Menu getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(14)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
