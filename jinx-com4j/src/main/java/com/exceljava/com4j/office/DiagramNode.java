package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0370-0000-0000-C000-000000000046}")
public interface DiagramNode extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoRelativeNodePosition parameter pos is set to 2</li><li>com.exceljava.com4j.office.MsoDiagramNodeType parameter nodeType is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNode(2, 1);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.DiagramNode
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(9)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {com.exceljava.com4j.office.MsoRelativeNodePosition.class, com.exceljava.com4j.office.MsoDiagramNodeType.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"2", "1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.DiagramNode addNode();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoDiagramNodeType parameter nodeType is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNode(pos, 1);
   * </code>
   * </p>
   * @param pos Optional parameter. Default value is 2
   * @return  Returns a value of type com.exceljava.com4j.office.DiagramNode
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(9)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {com.exceljava.com4j.office.MsoDiagramNodeType.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.DiagramNode addNode(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.MsoRelativeNodePosition pos);

  /**
   * @param pos Optional parameter. Default value is 2
   * @param nodeType Optional parameter. Default value is 1
   * @return  Returns a value of type com.exceljava.com4j.office.DiagramNode
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(9)
  com.exceljava.com4j.office.DiagramNode addNode(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.MsoRelativeNodePosition pos,
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoDiagramNodeType nodeType);


  /**
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(10)
  void delete();


  /**
   * @param targetNode Mandatory com.exceljava.com4j.office.DiagramNode parameter.
   * @param pos Mandatory com.exceljava.com4j.office.MsoRelativeNodePosition parameter.
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(11)
  void moveNode(
    com.exceljava.com4j.office.DiagramNode targetNode,
    com.exceljava.com4j.office.MsoRelativeNodePosition pos);


  /**
   * @param targetNode Mandatory com.exceljava.com4j.office.DiagramNode parameter.
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(12)
  void replaceNode(
    com.exceljava.com4j.office.DiagramNode targetNode);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter swapChildren is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * swapNode(targetNode, false);
   * </code>
   * </p>
   * @param targetNode Mandatory com.exceljava.com4j.office.DiagramNode parameter.
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {boolean.class}, nativeType = {NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL}, literal = {"false"})
  void swapNode(
    com.exceljava.com4j.office.DiagramNode targetNode);

  /**
   * @param targetNode Mandatory com.exceljava.com4j.office.DiagramNode parameter.
   * @param swapChildren Optional parameter. Default value is false
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(13)
  void swapNode(
    com.exceljava.com4j.office.DiagramNode targetNode,
    @Optional @DefaultValue("-1") boolean swapChildren);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoRelativeNodePosition parameter pos is set to 2</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * cloneNode(copyChildren, targetNode, 2);
   * </code>
   * </p>
   * @param copyChildren Mandatory boolean parameter.
   * @param targetNode Mandatory com.exceljava.com4j.office.DiagramNode parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.DiagramNode
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {com.exceljava.com4j.office.MsoRelativeNodePosition.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"2"})
  @ReturnValue(index=3)
  com.exceljava.com4j.office.DiagramNode cloneNode(
    boolean copyChildren,
    com.exceljava.com4j.office.DiagramNode targetNode);

  /**
   * @param copyChildren Mandatory boolean parameter.
   * @param targetNode Mandatory com.exceljava.com4j.office.DiagramNode parameter.
   * @param pos Optional parameter. Default value is 2
   * @return  Returns a value of type com.exceljava.com4j.office.DiagramNode
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.office.DiagramNode cloneNode(
    boolean copyChildren,
    com.exceljava.com4j.office.DiagramNode targetNode,
    @Optional @DefaultValue("2") com.exceljava.com4j.office.MsoRelativeNodePosition pos);


  /**
   * @param receivingNode Mandatory com.exceljava.com4j.office.DiagramNode parameter.
   */

  @DISPID(16) //= 0x10. The runtime will prefer the VTID if present
  @VTID(15)
  void transferChildren(
    com.exceljava.com4j.office.DiagramNode receivingNode);


  /**
   * @return  Returns a value of type com.exceljava.com4j.office.DiagramNode
   */

  @DISPID(17) //= 0x11. The runtime will prefer the VTID if present
  @VTID(16)
  com.exceljava.com4j.office.DiagramNode nextNode();


  /**
   * @return  Returns a value of type com.exceljava.com4j.office.DiagramNode
   */

  @DISPID(18) //= 0x12. The runtime will prefer the VTID if present
  @VTID(17)
  com.exceljava.com4j.office.DiagramNode prevNode();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(18)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Children"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.DiagramNodeChildren
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(19)
  com.exceljava.com4j.office.DiagramNodeChildren getChildren();


  @VTID(19)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.DiagramNodeChildren.class})
  com.exceljava.com4j.excel.DiagramNode getChildren(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "Shape"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(20)
  com.exceljava.com4j.office.Shape getShape();


  /**
   * <p>
   * Getter method for the COM property "Root"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.DiagramNode
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(21)
  com.exceljava.com4j.office.DiagramNode getRoot();


  /**
   * <p>
   * Getter method for the COM property "Diagram"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoDiagram
   */

  @DISPID(104) //= 0x68. The runtime will prefer the VTID if present
  @VTID(22)
  com.exceljava.com4j.office.IMsoDiagram getDiagram();


  /**
   * <p>
   * Getter method for the COM property "Layout"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoOrgChartLayoutType
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(23)
  com.exceljava.com4j.office.MsoOrgChartLayoutType getLayout();


  /**
   * <p>
   * Setter method for the COM property "Layout"
   * </p>
   * @param type Mandatory com.exceljava.com4j.office.MsoOrgChartLayoutType parameter.
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(24)
  void setLayout(
    com.exceljava.com4j.office.MsoOrgChartLayoutType type);


  /**
   * <p>
   * Getter method for the COM property "TextShape"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(106) //= 0x6a. The runtime will prefer the VTID if present
  @VTID(25)
  com.exceljava.com4j.office.Shape getTextShape();


  // Properties:
}
