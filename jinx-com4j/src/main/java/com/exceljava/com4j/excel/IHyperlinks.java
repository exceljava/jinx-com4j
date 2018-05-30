package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024430-0001-0000-C000-000000000046}")
public interface IHyperlinks extends Com4jObject,Iterable<Com4jObject> {
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
     * <li>java.lang.Object parameter subAddress is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter screenTip is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textToDisplay is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(anchor, address, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param anchor Mandatory com4j.Com4jObject parameter.
   * @param address Mandatory java.lang.String parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.Dispatch,index=5)
  com4j.Com4jObject add(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject anchor,
    java.lang.String address);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter screenTip is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textToDisplay is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(anchor, address, subAddress, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param anchor Mandatory com4j.Com4jObject parameter.
   * @param address Mandatory java.lang.String parameter.
   * @param subAddress Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.Dispatch,index=5)
  com4j.Com4jObject add(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject anchor,
    java.lang.String address,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subAddress);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter textToDisplay is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(anchor, address, subAddress, screenTip, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param anchor Mandatory com4j.Com4jObject parameter.
   * @param address Mandatory java.lang.String parameter.
   * @param subAddress Optional parameter. Default value is com4j.Variant.getMissing()
   * @param screenTip Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.Dispatch,index=5)
  com4j.Com4jObject add(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject anchor,
    java.lang.String address,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subAddress,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object screenTip);

  /**
   * @param anchor Mandatory com4j.Com4jObject parameter.
   * @param address Mandatory java.lang.String parameter.
   * @param subAddress Optional parameter. Default value is com4j.Variant.getMissing()
   * @param screenTip Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textToDisplay Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(10)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject add(
    @MarshalAs(NativeType.Dispatch) com4j.Com4jObject anchor,
    java.lang.String address,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object subAddress,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object screenTip,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textToDisplay);


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
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Hyperlink
   */

  @VTID(12)
  com.exceljava.com4j.excel.Hyperlink getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Hyperlink
   */

  @VTID(13)
  @DefaultMethod
  com.exceljava.com4j.excel.Hyperlink get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(14)
  java.util.Iterator<Com4jObject> iterator();

  /**
   */

  @VTID(15)
  void delete();


  // Properties:
}
