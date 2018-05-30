package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000CDB05-0000-0000-C000-000000000046}")
public interface _CustomXMLPart extends com.exceljava.com4j.office._IMsoDispObj {
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
   * Getter method for the COM property "DocumentElement"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLNode
   */

  @DISPID(1610809345) //= 0x60030001. The runtime will prefer the VTID if present
  @VTID(10)
  com.exceljava.com4j.office.CustomXMLNode getDocumentElement();


  /**
   * <p>
   * Getter method for the COM property "Id"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610809346) //= 0x60030002. The runtime will prefer the VTID if present
  @VTID(11)
  java.lang.String getId();


  /**
   * <p>
   * Getter method for the COM property "NamespaceURI"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610809347) //= 0x60030003. The runtime will prefer the VTID if present
  @VTID(12)
  java.lang.String getNamespaceURI();


  /**
   * <p>
   * Getter method for the COM property "SchemaCollection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office._CustomXMLSchemaCollection
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  com.exceljava.com4j.office._CustomXMLSchemaCollection getSchemaCollection();


  /**
   * <p>
   * Setter method for the COM property "SchemaCollection"
   * </p>
   * @param ppSchemaCollection Mandatory com.exceljava.com4j.office._CustomXMLSchemaCollection parameter.
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(14)
  void setSchemaCollection(
    com.exceljava.com4j.office._CustomXMLSchemaCollection ppSchemaCollection);


  /**
   * <p>
   * Getter method for the COM property "NamespaceManager"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLPrefixMappings
   */

  @DISPID(1610809350) //= 0x60030006. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.office.CustomXMLPrefixMappings getNamespaceManager();


  @VTID(15)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.CustomXMLPrefixMappings.class})
  com.exceljava.com4j.office.CustomXMLPrefixMapping getNamespaceManager(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "XML"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610809351) //= 0x60030007. The runtime will prefer the VTID if present
  @VTID(16)
  java.lang.String getXML();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter name is set to ""</li><li>java.lang.String parameter namespaceURI is set to ""</li><li>com.exceljava.com4j.office.CustomXMLNode parameter nextSibling is set to 0</li><li>com.exceljava.com4j.office.MsoCustomXMLNodeType parameter nodeType is set to 1</li><li>java.lang.String parameter nodeValue is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNode(parent, "", "", 0, 1, "");
   * </code>
   * </p>
   * @param parent Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   */

  @DISPID(1610809352) //= 0x60030008. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5}, javaType = {java.lang.String.class, java.lang.String.class, com.exceljava.com4j.office.CustomXMLNode.class, com.exceljava.com4j.office.MsoCustomXMLNodeType.class, java.lang.String.class}, nativeType = {NativeType.BSTR, NativeType.BSTR, NativeType.ComObject, NativeType.Int32, NativeType.BSTR}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_BSTR, Variant.Type.VT_I2, Variant.Type.VT_I4, Variant.Type.VT_BSTR}, literal = {"", "", "0", "1", ""})
  void addNode(
    com.exceljava.com4j.office.CustomXMLNode parent);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter namespaceURI is set to ""</li><li>com.exceljava.com4j.office.CustomXMLNode parameter nextSibling is set to 0</li><li>com.exceljava.com4j.office.MsoCustomXMLNodeType parameter nodeType is set to 1</li><li>java.lang.String parameter nodeValue is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNode(parent, name, "", 0, 1, "");
   * </code>
   * </p>
   * @param parent Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param name Optional parameter. Default value is ""
   */

  @DISPID(1610809352) //= 0x60030008. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5}, javaType = {java.lang.String.class, com.exceljava.com4j.office.CustomXMLNode.class, com.exceljava.com4j.office.MsoCustomXMLNodeType.class, java.lang.String.class}, nativeType = {NativeType.BSTR, NativeType.ComObject, NativeType.Int32, NativeType.BSTR}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_I2, Variant.Type.VT_I4, Variant.Type.VT_BSTR}, literal = {"", "0", "1", ""})
  void addNode(
    com.exceljava.com4j.office.CustomXMLNode parent,
    @Optional @DefaultValue("") java.lang.String name);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.CustomXMLNode parameter nextSibling is set to 0</li><li>com.exceljava.com4j.office.MsoCustomXMLNodeType parameter nodeType is set to 1</li><li>java.lang.String parameter nodeValue is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNode(parent, name, namespaceURI, 0, 1, "");
   * </code>
   * </p>
   * @param parent Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param name Optional parameter. Default value is ""
   * @param namespaceURI Optional parameter. Default value is ""
   */

  @DISPID(1610809352) //= 0x60030008. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5}, javaType = {com.exceljava.com4j.office.CustomXMLNode.class, com.exceljava.com4j.office.MsoCustomXMLNodeType.class, java.lang.String.class}, nativeType = {NativeType.ComObject, NativeType.Int32, NativeType.BSTR}, variantType = {Variant.Type.VT_I2, Variant.Type.VT_I4, Variant.Type.VT_BSTR}, literal = {"0", "1", ""})
  void addNode(
    com.exceljava.com4j.office.CustomXMLNode parent,
    @Optional @DefaultValue("") java.lang.String name,
    @Optional @DefaultValue("") java.lang.String namespaceURI);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoCustomXMLNodeType parameter nodeType is set to 1</li><li>java.lang.String parameter nodeValue is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNode(parent, name, namespaceURI, nextSibling, 1, "");
   * </code>
   * </p>
   * @param parent Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param name Optional parameter. Default value is ""
   * @param namespaceURI Optional parameter. Default value is ""
   * @param nextSibling Optional parameter. Default value is 0
   */

  @DISPID(1610809352) //= 0x60030008. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5}, javaType = {com.exceljava.com4j.office.MsoCustomXMLNodeType.class, java.lang.String.class}, nativeType = {NativeType.Int32, NativeType.BSTR}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_BSTR}, literal = {"1", ""})
  void addNode(
    com.exceljava.com4j.office.CustomXMLNode parent,
    @Optional @DefaultValue("") java.lang.String name,
    @Optional @DefaultValue("") java.lang.String namespaceURI,
    @Optional @DefaultValue("0") com.exceljava.com4j.office.CustomXMLNode nextSibling);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter nodeValue is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNode(parent, name, namespaceURI, nextSibling, nodeType, "");
   * </code>
   * </p>
   * @param parent Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param name Optional parameter. Default value is ""
   * @param namespaceURI Optional parameter. Default value is ""
   * @param nextSibling Optional parameter. Default value is 0
   * @param nodeType Optional parameter. Default value is 1
   */

  @DISPID(1610809352) //= 0x60030008. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5}, javaType = {java.lang.String.class}, nativeType = {NativeType.BSTR}, variantType = {Variant.Type.VT_BSTR}, literal = {""})
  void addNode(
    com.exceljava.com4j.office.CustomXMLNode parent,
    @Optional @DefaultValue("") java.lang.String name,
    @Optional @DefaultValue("") java.lang.String namespaceURI,
    @Optional @DefaultValue("0") com.exceljava.com4j.office.CustomXMLNode nextSibling,
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoCustomXMLNodeType nodeType);

  /**
   * @param parent Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param name Optional parameter. Default value is ""
   * @param namespaceURI Optional parameter. Default value is ""
   * @param nextSibling Optional parameter. Default value is 0
   * @param nodeType Optional parameter. Default value is 1
   * @param nodeValue Optional parameter. Default value is ""
   */

  @DISPID(1610809352) //= 0x60030008. The runtime will prefer the VTID if present
  @VTID(17)
  void addNode(
    com.exceljava.com4j.office.CustomXMLNode parent,
    @Optional @DefaultValue("") java.lang.String name,
    @Optional @DefaultValue("") java.lang.String namespaceURI,
    @Optional @DefaultValue("0") com.exceljava.com4j.office.CustomXMLNode nextSibling,
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoCustomXMLNodeType nodeType,
    @Optional @DefaultValue("") java.lang.String nodeValue);


  /**
   */

  @DISPID(1610809353) //= 0x60030009. The runtime will prefer the VTID if present
  @VTID(18)
  void delete();


  /**
   * @param filePath Mandatory java.lang.String parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(1610809354) //= 0x6003000a. The runtime will prefer the VTID if present
  @VTID(19)
  boolean load(
    java.lang.String filePath);


  /**
   * @param xml Mandatory java.lang.String parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(1610809355) //= 0x6003000b. The runtime will prefer the VTID if present
  @VTID(20)
  boolean loadXML(
    java.lang.String xml);


  /**
   * @param xPath Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLNodes
   */

  @DISPID(1610809356) //= 0x6003000c. The runtime will prefer the VTID if present
  @VTID(21)
  com.exceljava.com4j.office.CustomXMLNodes selectNodes(
    java.lang.String xPath);


  /**
   * @param xPath Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLNode
   */

  @DISPID(1610809357) //= 0x6003000d. The runtime will prefer the VTID if present
  @VTID(22)
  com.exceljava.com4j.office.CustomXMLNode selectSingleNode(
    java.lang.String xPath);


  /**
   * <p>
   * Getter method for the COM property "Errors"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLValidationErrors
   */

  @DISPID(1610809358) //= 0x6003000e. The runtime will prefer the VTID if present
  @VTID(23)
  com.exceljava.com4j.office.CustomXMLValidationErrors getErrors();


  @VTID(23)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.CustomXMLValidationErrors.class})
  com.exceljava.com4j.office.CustomXMLValidationError getErrors(
    int index);

  /**
   * <p>
   * Getter method for the COM property "BuiltIn"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610809359) //= 0x6003000f. The runtime will prefer the VTID if present
  @VTID(24)
  boolean getBuiltIn();


  // Properties:
}
