package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024474-0001-0000-C000-000000000046}")
public interface IListRows extends Com4jObject,Iterable<Com4jObject> {
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
     * <li>java.lang.Object parameter position is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Add(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ListRow
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.ListRow _Add();

  /**
   * @param position Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ListRow
   */

  @VTID(10)
  com.exceljava.com4j.excel.ListRow _Add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object position);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ListRow
   */

  @VTID(11)
  @DefaultMethod
  com.exceljava.com4j.excel.ListRow get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(12)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ListRow
   */

  @VTID(13)
  com.exceljava.com4j.excel.ListRow getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(14)
  int getCount();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter position is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter alwaysInsert is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ListRow
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.ListRow add();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter alwaysInsert is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(position, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param position Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ListRow
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.ListRow add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object position);

  /**
   * @param position Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alwaysInsert Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.ListRow
   */

  @VTID(15)
  com.exceljava.com4j.excel.ListRow add(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object position,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object alwaysInsert);


  // Properties:
}
