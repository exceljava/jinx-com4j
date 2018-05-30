package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000CDB04-0000-0000-C000-000000000046}")
public interface CustomXMLNode extends com.exceljava.com4j.office._IMsoDispObj {
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
   * Getter method for the COM property "Attributes"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLNodes
   */

  @DISPID(1610809345) //= 0x60030001. The runtime will prefer the VTID if present
  @VTID(10)
  com.exceljava.com4j.office.CustomXMLNodes getAttributes();


  @VTID(10)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.CustomXMLNodes.class})
  com.exceljava.com4j.office.CustomXMLNode getAttributes(
    int index);

  /**
   * <p>
   * Getter method for the COM property "BaseName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610809346) //= 0x60030002. The runtime will prefer the VTID if present
  @VTID(11)
  java.lang.String getBaseName();


  /**
   * <p>
   * Getter method for the COM property "ChildNodes"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLNodes
   */

  @DISPID(1610809347) //= 0x60030003. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.CustomXMLNodes getChildNodes();


  @VTID(12)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.CustomXMLNodes.class})
  com.exceljava.com4j.office.CustomXMLNode getChildNodes(
    int index);

  /**
   * <p>
   * Getter method for the COM property "FirstChild"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLNode
   */

  @DISPID(1610809348) //= 0x60030004. The runtime will prefer the VTID if present
  @VTID(13)
  com.exceljava.com4j.office.CustomXMLNode getFirstChild();


  /**
   * <p>
   * Getter method for the COM property "LastChild"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLNode
   */

  @DISPID(1610809349) //= 0x60030005. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.office.CustomXMLNode getLastChild();


  /**
   * <p>
   * Getter method for the COM property "NamespaceURI"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610809350) //= 0x60030006. The runtime will prefer the VTID if present
  @VTID(15)
  java.lang.String getNamespaceURI();


  /**
   * <p>
   * Getter method for the COM property "NextSibling"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLNode
   */

  @DISPID(1610809351) //= 0x60030007. The runtime will prefer the VTID if present
  @VTID(16)
  com.exceljava.com4j.office.CustomXMLNode getNextSibling();


  /**
   * <p>
   * Getter method for the COM property "NodeType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoCustomXMLNodeType
   */

  @DISPID(1610809352) //= 0x60030008. The runtime will prefer the VTID if present
  @VTID(17)
  com.exceljava.com4j.office.MsoCustomXMLNodeType getNodeType();


  /**
   * <p>
   * Getter method for the COM property "NodeValue"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610809353) //= 0x60030009. The runtime will prefer the VTID if present
  @VTID(18)
  java.lang.String getNodeValue();


  /**
   * <p>
   * Setter method for the COM property "NodeValue"
   * </p>
   * @param pbstrNodeValue Mandatory java.lang.String parameter.
   */

  @DISPID(1610809353) //= 0x60030009. The runtime will prefer the VTID if present
  @VTID(19)
  void setNodeValue(
    java.lang.String pbstrNodeValue);


  /**
   * <p>
   * Getter method for the COM property "OwnerDocument"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610809355) //= 0x6003000b. The runtime will prefer the VTID if present
  @VTID(20)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getOwnerDocument();


  /**
   * <p>
   * Getter method for the COM property "OwnerPart"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office._CustomXMLPart
   */

  @DISPID(1610809356) //= 0x6003000c. The runtime will prefer the VTID if present
  @VTID(21)
  com.exceljava.com4j.office._CustomXMLPart getOwnerPart();


  /**
   * <p>
   * Getter method for the COM property "PreviousSibling"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLNode
   */

  @DISPID(1610809357) //= 0x6003000d. The runtime will prefer the VTID if present
  @VTID(22)
  com.exceljava.com4j.office.CustomXMLNode getPreviousSibling();


  /**
   * <p>
   * Getter method for the COM property "ParentNode"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLNode
   */

  @DISPID(1610809358) //= 0x6003000e. The runtime will prefer the VTID if present
  @VTID(23)
  com.exceljava.com4j.office.CustomXMLNode getParentNode();


  /**
   * <p>
   * Getter method for the COM property "Text"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610809359) //= 0x6003000f. The runtime will prefer the VTID if present
  @VTID(24)
  java.lang.String getText();


  /**
   * <p>
   * Setter method for the COM property "Text"
   * </p>
   * @param pbstrText Mandatory java.lang.String parameter.
   */

  @DISPID(1610809359) //= 0x6003000f. The runtime will prefer the VTID if present
  @VTID(25)
  void setText(
    java.lang.String pbstrText);


  /**
   * <p>
   * Getter method for the COM property "XPath"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610809361) //= 0x60030011. The runtime will prefer the VTID if present
  @VTID(26)
  java.lang.String getXPath();


  /**
   * <p>
   * Getter method for the COM property "XML"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610809362) //= 0x60030012. The runtime will prefer the VTID if present
  @VTID(27)
  java.lang.String getXML();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter name is set to ""</li><li>java.lang.String parameter namespaceURI is set to ""</li><li>com.exceljava.com4j.office.MsoCustomXMLNodeType parameter nodeType is set to 1</li><li>java.lang.String parameter nodeValue is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * appendChildNode("", "", 1, "");
   * </code>
   * </p>
   */

  @DISPID(1610809363) //= 0x60030013. The runtime will prefer the VTID if present
  @VTID(28)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.String.class, java.lang.String.class, com.exceljava.com4j.office.MsoCustomXMLNodeType.class, java.lang.String.class}, nativeType = {NativeType.BSTR, NativeType.BSTR, NativeType.Int32, NativeType.BSTR}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_BSTR, Variant.Type.VT_I4, Variant.Type.VT_BSTR}, literal = {"", "", "1", ""})
  void appendChildNode();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter namespaceURI is set to ""</li><li>com.exceljava.com4j.office.MsoCustomXMLNodeType parameter nodeType is set to 1</li><li>java.lang.String parameter nodeValue is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * appendChildNode(name, "", 1, "");
   * </code>
   * </p>
   * @param name Optional parameter. Default value is ""
   */

  @DISPID(1610809363) //= 0x60030013. The runtime will prefer the VTID if present
  @VTID(28)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.String.class, com.exceljava.com4j.office.MsoCustomXMLNodeType.class, java.lang.String.class}, nativeType = {NativeType.BSTR, NativeType.Int32, NativeType.BSTR}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_I4, Variant.Type.VT_BSTR}, literal = {"", "1", ""})
  void appendChildNode(
    @Optional @DefaultValue("") java.lang.String name);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoCustomXMLNodeType parameter nodeType is set to 1</li><li>java.lang.String parameter nodeValue is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * appendChildNode(name, namespaceURI, 1, "");
   * </code>
   * </p>
   * @param name Optional parameter. Default value is ""
   * @param namespaceURI Optional parameter. Default value is ""
   */

  @DISPID(1610809363) //= 0x60030013. The runtime will prefer the VTID if present
  @VTID(28)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {com.exceljava.com4j.office.MsoCustomXMLNodeType.class, java.lang.String.class}, nativeType = {NativeType.Int32, NativeType.BSTR}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_BSTR}, literal = {"1", ""})
  void appendChildNode(
    @Optional @DefaultValue("") java.lang.String name,
    @Optional @DefaultValue("") java.lang.String namespaceURI);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter nodeValue is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * appendChildNode(name, namespaceURI, nodeType, "");
   * </code>
   * </p>
   * @param name Optional parameter. Default value is ""
   * @param namespaceURI Optional parameter. Default value is ""
   * @param nodeType Optional parameter. Default value is 1
   */

  @DISPID(1610809363) //= 0x60030013. The runtime will prefer the VTID if present
  @VTID(28)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.String.class}, nativeType = {NativeType.BSTR}, variantType = {Variant.Type.VT_BSTR}, literal = {""})
  void appendChildNode(
    @Optional @DefaultValue("") java.lang.String name,
    @Optional @DefaultValue("") java.lang.String namespaceURI,
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoCustomXMLNodeType nodeType);

  /**
   * @param name Optional parameter. Default value is ""
   * @param namespaceURI Optional parameter. Default value is ""
   * @param nodeType Optional parameter. Default value is 1
   * @param nodeValue Optional parameter. Default value is ""
   */

  @DISPID(1610809363) //= 0x60030013. The runtime will prefer the VTID if present
  @VTID(28)
  void appendChildNode(
    @Optional @DefaultValue("") java.lang.String name,
    @Optional @DefaultValue("") java.lang.String namespaceURI,
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoCustomXMLNodeType nodeType,
    @Optional @DefaultValue("") java.lang.String nodeValue);


  /**
   * @param xml Mandatory java.lang.String parameter.
   */

  @DISPID(1610809364) //= 0x60030014. The runtime will prefer the VTID if present
  @VTID(29)
  void appendChildSubtree(
    java.lang.String xml);


  /**
   */

  @DISPID(1610809365) //= 0x60030015. The runtime will prefer the VTID if present
  @VTID(30)
  void delete();


  /**
   * @return  Returns a value of type boolean
   */

  @DISPID(1610809366) //= 0x60030016. The runtime will prefer the VTID if present
  @VTID(31)
  boolean hasChildNodes();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter name is set to ""</li><li>java.lang.String parameter namespaceURI is set to ""</li><li>com.exceljava.com4j.office.MsoCustomXMLNodeType parameter nodeType is set to 1</li><li>java.lang.String parameter nodeValue is set to ""</li><li>com.exceljava.com4j.office.CustomXMLNode parameter nextSibling is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insertNodeBefore("", "", 1, "", 0);
   * </code>
   * </p>
   */

  @DISPID(1610809367) //= 0x60030017. The runtime will prefer the VTID if present
  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {java.lang.String.class, java.lang.String.class, com.exceljava.com4j.office.MsoCustomXMLNodeType.class, java.lang.String.class, com.exceljava.com4j.office.CustomXMLNode.class}, nativeType = {NativeType.BSTR, NativeType.BSTR, NativeType.Int32, NativeType.BSTR, NativeType.ComObject}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_BSTR, Variant.Type.VT_I4, Variant.Type.VT_BSTR, Variant.Type.VT_I2}, literal = {"", "", "1", "", "0"})
  void insertNodeBefore();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter namespaceURI is set to ""</li><li>com.exceljava.com4j.office.MsoCustomXMLNodeType parameter nodeType is set to 1</li><li>java.lang.String parameter nodeValue is set to ""</li><li>com.exceljava.com4j.office.CustomXMLNode parameter nextSibling is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insertNodeBefore(name, "", 1, "", 0);
   * </code>
   * </p>
   * @param name Optional parameter. Default value is ""
   */

  @DISPID(1610809367) //= 0x60030017. The runtime will prefer the VTID if present
  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.String.class, com.exceljava.com4j.office.MsoCustomXMLNodeType.class, java.lang.String.class, com.exceljava.com4j.office.CustomXMLNode.class}, nativeType = {NativeType.BSTR, NativeType.Int32, NativeType.BSTR, NativeType.ComObject}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_I4, Variant.Type.VT_BSTR, Variant.Type.VT_I2}, literal = {"", "1", "", "0"})
  void insertNodeBefore(
    @Optional @DefaultValue("") java.lang.String name);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoCustomXMLNodeType parameter nodeType is set to 1</li><li>java.lang.String parameter nodeValue is set to ""</li><li>com.exceljava.com4j.office.CustomXMLNode parameter nextSibling is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insertNodeBefore(name, namespaceURI, 1, "", 0);
   * </code>
   * </p>
   * @param name Optional parameter. Default value is ""
   * @param namespaceURI Optional parameter. Default value is ""
   */

  @DISPID(1610809367) //= 0x60030017. The runtime will prefer the VTID if present
  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {com.exceljava.com4j.office.MsoCustomXMLNodeType.class, java.lang.String.class, com.exceljava.com4j.office.CustomXMLNode.class}, nativeType = {NativeType.Int32, NativeType.BSTR, NativeType.ComObject}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_BSTR, Variant.Type.VT_I2}, literal = {"1", "", "0"})
  void insertNodeBefore(
    @Optional @DefaultValue("") java.lang.String name,
    @Optional @DefaultValue("") java.lang.String namespaceURI);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter nodeValue is set to ""</li><li>com.exceljava.com4j.office.CustomXMLNode parameter nextSibling is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insertNodeBefore(name, namespaceURI, nodeType, "", 0);
   * </code>
   * </p>
   * @param name Optional parameter. Default value is ""
   * @param namespaceURI Optional parameter. Default value is ""
   * @param nodeType Optional parameter. Default value is 1
   */

  @DISPID(1610809367) //= 0x60030017. The runtime will prefer the VTID if present
  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.String.class, com.exceljava.com4j.office.CustomXMLNode.class}, nativeType = {NativeType.BSTR, NativeType.ComObject}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_I2}, literal = {"", "0"})
  void insertNodeBefore(
    @Optional @DefaultValue("") java.lang.String name,
    @Optional @DefaultValue("") java.lang.String namespaceURI,
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoCustomXMLNodeType nodeType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.CustomXMLNode parameter nextSibling is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insertNodeBefore(name, namespaceURI, nodeType, nodeValue, 0);
   * </code>
   * </p>
   * @param name Optional parameter. Default value is ""
   * @param namespaceURI Optional parameter. Default value is ""
   * @param nodeType Optional parameter. Default value is 1
   * @param nodeValue Optional parameter. Default value is ""
   */

  @DISPID(1610809367) //= 0x60030017. The runtime will prefer the VTID if present
  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {com.exceljava.com4j.office.CustomXMLNode.class}, nativeType = {NativeType.ComObject}, variantType = {Variant.Type.VT_I2}, literal = {"0"})
  void insertNodeBefore(
    @Optional @DefaultValue("") java.lang.String name,
    @Optional @DefaultValue("") java.lang.String namespaceURI,
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoCustomXMLNodeType nodeType,
    @Optional @DefaultValue("") java.lang.String nodeValue);

  /**
   * @param name Optional parameter. Default value is ""
   * @param namespaceURI Optional parameter. Default value is ""
   * @param nodeType Optional parameter. Default value is 1
   * @param nodeValue Optional parameter. Default value is ""
   * @param nextSibling Optional parameter. Default value is 0
   */

  @DISPID(1610809367) //= 0x60030017. The runtime will prefer the VTID if present
  @VTID(32)
  void insertNodeBefore(
    @Optional @DefaultValue("") java.lang.String name,
    @Optional @DefaultValue("") java.lang.String namespaceURI,
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoCustomXMLNodeType nodeType,
    @Optional @DefaultValue("") java.lang.String nodeValue,
    @Optional @DefaultValue("0") com.exceljava.com4j.office.CustomXMLNode nextSibling);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.CustomXMLNode parameter nextSibling is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insertSubtreeBefore(xml, 0);
   * </code>
   * </p>
   * @param xml Mandatory java.lang.String parameter.
   */

  @DISPID(1610809368) //= 0x60030018. The runtime will prefer the VTID if present
  @VTID(33)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {com.exceljava.com4j.office.CustomXMLNode.class}, nativeType = {NativeType.ComObject}, variantType = {Variant.Type.VT_I2}, literal = {"0"})
  void insertSubtreeBefore(
    java.lang.String xml);

  /**
   * @param xml Mandatory java.lang.String parameter.
   * @param nextSibling Optional parameter. Default value is 0
   */

  @DISPID(1610809368) //= 0x60030018. The runtime will prefer the VTID if present
  @VTID(33)
  void insertSubtreeBefore(
    java.lang.String xml,
    @Optional @DefaultValue("0") com.exceljava.com4j.office.CustomXMLNode nextSibling);


  /**
   * @param child Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   */

  @DISPID(1610809369) //= 0x60030019. The runtime will prefer the VTID if present
  @VTID(34)
  void removeChild(
    com.exceljava.com4j.office.CustomXMLNode child);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter name is set to ""</li><li>java.lang.String parameter namespaceURI is set to ""</li><li>com.exceljava.com4j.office.MsoCustomXMLNodeType parameter nodeType is set to 1</li><li>java.lang.String parameter nodeValue is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * replaceChildNode(oldNode, "", "", 1, "");
   * </code>
   * </p>
   * @param oldNode Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   */

  @DISPID(1610809370) //= 0x6003001a. The runtime will prefer the VTID if present
  @VTID(35)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.String.class, java.lang.String.class, com.exceljava.com4j.office.MsoCustomXMLNodeType.class, java.lang.String.class}, nativeType = {NativeType.BSTR, NativeType.BSTR, NativeType.Int32, NativeType.BSTR}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_BSTR, Variant.Type.VT_I4, Variant.Type.VT_BSTR}, literal = {"", "", "1", ""})
  void replaceChildNode(
    com.exceljava.com4j.office.CustomXMLNode oldNode);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter namespaceURI is set to ""</li><li>com.exceljava.com4j.office.MsoCustomXMLNodeType parameter nodeType is set to 1</li><li>java.lang.String parameter nodeValue is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * replaceChildNode(oldNode, name, "", 1, "");
   * </code>
   * </p>
   * @param oldNode Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param name Optional parameter. Default value is ""
   */

  @DISPID(1610809370) //= 0x6003001a. The runtime will prefer the VTID if present
  @VTID(35)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {java.lang.String.class, com.exceljava.com4j.office.MsoCustomXMLNodeType.class, java.lang.String.class}, nativeType = {NativeType.BSTR, NativeType.Int32, NativeType.BSTR}, variantType = {Variant.Type.VT_BSTR, Variant.Type.VT_I4, Variant.Type.VT_BSTR}, literal = {"", "1", ""})
  void replaceChildNode(
    com.exceljava.com4j.office.CustomXMLNode oldNode,
    @Optional @DefaultValue("") java.lang.String name);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoCustomXMLNodeType parameter nodeType is set to 1</li><li>java.lang.String parameter nodeValue is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * replaceChildNode(oldNode, name, namespaceURI, 1, "");
   * </code>
   * </p>
   * @param oldNode Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param name Optional parameter. Default value is ""
   * @param namespaceURI Optional parameter. Default value is ""
   */

  @DISPID(1610809370) //= 0x6003001a. The runtime will prefer the VTID if present
  @VTID(35)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {com.exceljava.com4j.office.MsoCustomXMLNodeType.class, java.lang.String.class}, nativeType = {NativeType.Int32, NativeType.BSTR}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_BSTR}, literal = {"1", ""})
  void replaceChildNode(
    com.exceljava.com4j.office.CustomXMLNode oldNode,
    @Optional @DefaultValue("") java.lang.String name,
    @Optional @DefaultValue("") java.lang.String namespaceURI);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.String parameter nodeValue is set to ""</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * replaceChildNode(oldNode, name, namespaceURI, nodeType, "");
   * </code>
   * </p>
   * @param oldNode Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param name Optional parameter. Default value is ""
   * @param namespaceURI Optional parameter. Default value is ""
   * @param nodeType Optional parameter. Default value is 1
   */

  @DISPID(1610809370) //= 0x6003001a. The runtime will prefer the VTID if present
  @VTID(35)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.String.class}, nativeType = {NativeType.BSTR}, variantType = {Variant.Type.VT_BSTR}, literal = {""})
  void replaceChildNode(
    com.exceljava.com4j.office.CustomXMLNode oldNode,
    @Optional @DefaultValue("") java.lang.String name,
    @Optional @DefaultValue("") java.lang.String namespaceURI,
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoCustomXMLNodeType nodeType);

  /**
   * @param oldNode Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   * @param name Optional parameter. Default value is ""
   * @param namespaceURI Optional parameter. Default value is ""
   * @param nodeType Optional parameter. Default value is 1
   * @param nodeValue Optional parameter. Default value is ""
   */

  @DISPID(1610809370) //= 0x6003001a. The runtime will prefer the VTID if present
  @VTID(35)
  void replaceChildNode(
    com.exceljava.com4j.office.CustomXMLNode oldNode,
    @Optional @DefaultValue("") java.lang.String name,
    @Optional @DefaultValue("") java.lang.String namespaceURI,
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoCustomXMLNodeType nodeType,
    @Optional @DefaultValue("") java.lang.String nodeValue);


  /**
   * @param xml Mandatory java.lang.String parameter.
   * @param oldNode Mandatory com.exceljava.com4j.office.CustomXMLNode parameter.
   */

  @DISPID(1610809371) //= 0x6003001b. The runtime will prefer the VTID if present
  @VTID(36)
  void replaceChildSubtree(
    java.lang.String xml,
    com.exceljava.com4j.office.CustomXMLNode oldNode);


  /**
   * @param xPath Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLNodes
   */

  @DISPID(1610809372) //= 0x6003001c. The runtime will prefer the VTID if present
  @VTID(37)
  com.exceljava.com4j.office.CustomXMLNodes selectNodes(
    java.lang.String xPath);


  /**
   * @param xPath Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.CustomXMLNode
   */

  @DISPID(1610809373) //= 0x6003001d. The runtime will prefer the VTID if present
  @VTID(38)
  com.exceljava.com4j.office.CustomXMLNode selectSingleNode(
    java.lang.String xPath);


  // Properties:
}
