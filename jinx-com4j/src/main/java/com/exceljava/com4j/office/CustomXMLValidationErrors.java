package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000CDB0F-0000-0000-C000-000000000046}")
public interface CustomXMLValidationErrors extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610809344) //= 0x60030000. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610809345) //= 0x60030001. The runtime will prefer the VTID if present
  @VTID(10)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLValidationError
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(11)
  @DefaultMethod
  com.exceljava.com4j.office.CustomXMLValidationError getItem(
    int index);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter errorText is set to ""</li><li>boolean parameter clearedOnUpdate is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(node, errorName, "", false);
   * </code>
   * </p>
   * @param node Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param errorName Mandatory java.lang.String parameter.
   */

  @DISPID(1610809347) //= 0x60030003. The runtime will prefer the VTID if present
  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.String.class, boolean.class}, nativeType = {NativeType.BSTR, NativeType.VariantBool}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_BOOL}, literal = {"", "false"})
  void add(
    com.exceljava.com4j.office.CustomXMLNode node,
    java.lang.String errorName);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter clearedOnUpdate is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(node, errorName, errorText, false);
   * </code>
   * </p>
   * @param node Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param errorName Mandatory java.lang.String parameter.
   * @param errorText Optional parameter. Default value is ""
   */

  @DISPID(1610809347) //= 0x60030003. The runtime will prefer the VTID if present
  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {boolean.class}, nativeType = {NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL}, literal = {"false"})
  void add(
    com.exceljava.com4j.office.CustomXMLNode node,
    java.lang.String errorName,
    @Optional @DefaultValue("") java.lang.String errorText);

  /**
   * @param node Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param errorName Mandatory java.lang.String parameter.
   * @param errorText Optional parameter. Default value is ""
   * @param clearedOnUpdate Optional parameter. Default value is false
   */

  @DISPID(1610809347) //= 0x60030003. The runtime will prefer the VTID if present
  @VTID(12)
  void add(
    com.exceljava.com4j.office.CustomXMLNode node,
    java.lang.String errorName,
    @Optional @DefaultValue("") java.lang.String errorText,
    @Optional @DefaultValue("-1") boolean clearedOnUpdate);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(13)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
