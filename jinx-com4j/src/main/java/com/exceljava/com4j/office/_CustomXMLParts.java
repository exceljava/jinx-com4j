package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000CDB09-0000-0000-C000-000000000046}")
public interface _CustomXMLParts extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
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
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.office._CustomXMLPart
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(11)
  @DefaultMethod
  com.exceljava.com4j.office._CustomXMLPart getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter xml is set to ""</li><li>java.lang.Object parameter schemaCollection is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add("", com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office._CustomXMLPart
   */

  @DISPID(1610809347) //= 0x60030003. The runtime will prefer the VTID if present
  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.String.class, java.lang.Object.class}, nativeType = {NativeType.BSTR, NativeType.VARIANT}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_ERROR}, literal = {"", "80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office._CustomXMLPart add();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter schemaCollection is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(xml, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param xml Optional parameter. Default value is ""
   * @return  Returns a value of type com.exceljava.com4j.office._CustomXMLPart
   */

  @DISPID(1610809347) //= 0x60030003. The runtime will prefer the VTID if present
  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office._CustomXMLPart add(
    @Optional @DefaultValue("") java.lang.String xml);

  /**
   * @param xml Optional parameter. Default value is ""
   * @param schemaCollection Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office._CustomXMLPart
   */

  @DISPID(1610809347) //= 0x60030003. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office._CustomXMLPart add(
    @Optional @DefaultValue("") java.lang.String xml,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object schemaCollection);


  /**
   * @param id Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.office._CustomXMLPart
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  com.exceljava.com4j.office._CustomXMLPart selectByID(
    java.lang.String id);


  /**
   * @param namespaceURI Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.office._CustomXMLParts
   */

  @DISPID(1610809349) //= 0x60030005. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.office._CustomXMLParts selectByNamespace(
    java.lang.String namespaceURI);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(15)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
