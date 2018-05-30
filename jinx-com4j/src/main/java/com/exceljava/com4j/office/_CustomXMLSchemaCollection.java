package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000CDB02-0000-0000-C000-000000000046}")
public interface _CustomXMLSchemaCollection extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
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
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLSchema
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(11)
  @DefaultMethod
  com.exceljava.com4j.office.CustomXMLSchema getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "NamespaceURI"
   * </p>
   * @param index Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610809347) //= 0x60030003. The runtime will prefer the VTID if present
  @VTID(12)
  java.lang.String getNamespaceURI(
    int index);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter namespaceURI is set to ""</li><li>java.lang.String parameter alias is set to ""</li><li>java.lang.String parameter fileName is set to ""</li><li>boolean parameter installForAllUsers is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add("", "", "", false);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLSchema
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {4}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.String.class, java.lang.String.class, java.lang.String.class, boolean.class}, nativeType = {NativeType.BSTR, NativeType.BSTR, NativeType.BSTR, NativeType.VariantBool}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_BSTR, Variant.Type.VT_BSTR, Variant.Type.VT_BOOL}, literal = {"", "", "", "false"})
  @ReturnValue(index=4)
  com.exceljava.com4j.office.CustomXMLSchema add();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter alias is set to ""</li><li>java.lang.String parameter fileName is set to ""</li><li>boolean parameter installForAllUsers is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(namespaceURI, "", "", false);
   * </code>
   * </p>
   * @param namespaceURI Optional parameter. Default value is ""
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLSchema
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {java.lang.String.class, java.lang.String.class, boolean.class}, nativeType = {NativeType.BSTR, NativeType.BSTR, NativeType.VariantBool}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_BSTR, Variant.Type.VT_BOOL}, literal = {"", "", "false"})
  @ReturnValue(index=4)
  com.exceljava.com4j.office.CustomXMLSchema add(
    @Optional @DefaultValue("") java.lang.String namespaceURI);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter fileName is set to ""</li><li>boolean parameter installForAllUsers is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(namespaceURI, alias, "", false);
   * </code>
   * </p>
   * @param namespaceURI Optional parameter. Default value is ""
   * @param alias Optional parameter. Default value is ""
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLSchema
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.String.class, boolean.class}, nativeType = {NativeType.BSTR, NativeType.VariantBool}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_BOOL}, literal = {"", "false"})
  @ReturnValue(index=4)
  com.exceljava.com4j.office.CustomXMLSchema add(
    @Optional @DefaultValue("") java.lang.String namespaceURI,
    @Optional @DefaultValue("") java.lang.String alias);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter installForAllUsers is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add(namespaceURI, alias, fileName, false);
   * </code>
   * </p>
   * @param namespaceURI Optional parameter. Default value is ""
   * @param alias Optional parameter. Default value is ""
   * @param fileName Optional parameter. Default value is ""
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLSchema
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {boolean.class}, nativeType = {NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL}, literal = {"false"})
  @ReturnValue(index=4)
  com.exceljava.com4j.office.CustomXMLSchema add(
    @Optional @DefaultValue("") java.lang.String namespaceURI,
    @Optional @DefaultValue("") java.lang.String alias,
    @Optional @DefaultValue("") java.lang.String fileName);

  /**
   * @param namespaceURI Optional parameter. Default value is ""
   * @param alias Optional parameter. Default value is ""
   * @param fileName Optional parameter. Default value is ""
   * @param installForAllUsers Optional parameter. Default value is false
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLSchema
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  com.exceljava.com4j.office.CustomXMLSchema add(
    @Optional @DefaultValue("") java.lang.String namespaceURI,
    @Optional @DefaultValue("") java.lang.String alias,
    @Optional @DefaultValue("") java.lang.String fileName,
    @Optional @DefaultValue("0") boolean installForAllUsers);


  /**
   * @param schemaCollection Mandatory com.exceljava.com4j.office._CustomXMLSchemaCollection parameter.
   */

  @DISPID(1610809349) //= 0x60030005. The runtime will prefer the VTID if present
  @VTID(14)
  void addCollection(
    com.exceljava.com4j.office._CustomXMLSchemaCollection schemaCollection);


  /**
   * @return  Returns a value of type boolean
   */

  @DISPID(1610809350) //= 0x60030006. The runtime will prefer the VTID if present
  @VTID(15)
  boolean validate();


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(16)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
