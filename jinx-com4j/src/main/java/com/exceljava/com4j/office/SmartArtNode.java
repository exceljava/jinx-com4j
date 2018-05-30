package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C03C8-0000-0000-C000-000000000046}")
public interface SmartArtNode extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoSmartArtNodePosition parameter position is set to 1</li><li>com.exceljava.com4j.office.MsoSmartArtNodeType parameter type is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNode(1, 1);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SmartArtNode
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {com.exceljava.com4j.office.MsoSmartArtNodePosition.class, com.exceljava.com4j.office.MsoSmartArtNodeType.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.SmartArtNode addNode();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoSmartArtNodeType parameter type is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNode(position, 1);
   * </code>
   * </p>
   * @param position Optional parameter. Default value is 1
   * @return  Returns a value of type com.exceljava.com4j.office.SmartArtNode
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {com.exceljava.com4j.office.MsoSmartArtNodeType.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.SmartArtNode addNode(
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoSmartArtNodePosition position);

  /**
   * @param position Optional parameter. Default value is 1
   * @param type Optional parameter. Default value is 1
   * @return  Returns a value of type com.exceljava.com4j.office.SmartArtNode
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(10)
  com.exceljava.com4j.office.SmartArtNode addNode(
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoSmartArtNodePosition position,
    @Optional @DefaultValue("1") com.exceljava.com4j.office.MsoSmartArtNodeType type);


  /**
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(11)
  void delete();


  /**
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(12)
  void promote();


  /**
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(13)
  void demote();


  /**
   * <p>
   * Getter method for the COM property "OrgChartLayout"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoOrgChartLayoutType
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.office.MsoOrgChartLayoutType getOrgChartLayout();


  /**
   * <p>
   * Setter method for the COM property "OrgChartLayout"
   * </p>
   * @param type Mandatory com.exceljava.com4j.office.MsoOrgChartLayoutType parameter.
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(15)
  void setOrgChartLayout(
    com.exceljava.com4j.office.MsoOrgChartLayoutType type);


  /**
   * <p>
   * Getter method for the COM property "Shapes"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.ShapeRange
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(16)
  com.exceljava.com4j.office.ShapeRange getShapes();


  @VTID(16)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.ShapeRange.class})
  com.exceljava.com4j.office.Shape getShapes(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "TextFrame2"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.TextFrame2
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(17)
  com.exceljava.com4j.office.TextFrame2 getTextFrame2();


  /**
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(18)
  void larger();


  /**
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(19)
  void smaller();


  /**
   * <p>
   * Getter method for the COM property "Level"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(20)
  int getLevel();


  /**
   * <p>
   * Getter method for the COM property "Hidden"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(21)
  com.exceljava.com4j.office.MsoTriState getHidden();


  /**
   * <p>
   * Getter method for the COM property "Nodes"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SmartArtNodes
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(22)
  com.exceljava.com4j.office.SmartArtNodes getNodes();


  @VTID(22)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.SmartArtNodes.class})
  com.exceljava.com4j.office.SmartArtNode getNodes(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "ParentNode"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SmartArtNode
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(23)
  com.exceljava.com4j.office.SmartArtNode getParentNode();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoSmartArtNodeType
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(24)
  com.exceljava.com4j.office.MsoSmartArtNodeType getType();


  /**
   */

  @DISPID(16) //= 0x10. The runtime will prefer the VTID if present
  @VTID(25)
  void reorderUp();


  /**
   */

  @DISPID(17) //= 0x11. The runtime will prefer the VTID if present
  @VTID(26)
  void reorderDown();


  // Properties:
}
