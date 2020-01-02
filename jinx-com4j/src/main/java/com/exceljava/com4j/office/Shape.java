package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C031C-0000-0000-C000-000000000046}")
public interface Shape extends com.exceljava.com4j.office._IMsoDispObj {
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
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(10)
  void apply();


  /**
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(11)
  void delete();


  /**
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.Shape duplicate();


  /**
   * @param flipCmd Mandatory com.exceljava.com4j.office.MsoFlipCmd parameter.
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(13)
  void flip(
    com.exceljava.com4j.office.MsoFlipCmd flipCmd);


  /**
   * @param increment Mandatory float parameter.
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(14)
  void incrementLeft(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(15)
  void incrementRotation(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @DISPID(16) //= 0x10. The runtime will prefer the VTID if present
  @VTID(16)
  void incrementTop(
    float increment);


  /**
   */

  @DISPID(17) //= 0x11. The runtime will prefer the VTID if present
  @VTID(17)
  void pickUp();


  /**
   */

  @DISPID(18) //= 0x12. The runtime will prefer the VTID if present
  @VTID(18)
  void rerouteConnections();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoScaleFrom parameter fScale is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * scaleHeight(factor, relativeToOriginalSize, 0);
   * </code>
   * </p>
   * @param factor Mandatory float parameter.
   * @param relativeToOriginalSize Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(19) //= 0x13. The runtime will prefer the VTID if present
  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {com.exceljava.com4j.office.MsoScaleFrom.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  void scaleHeight(
    float factor,
    com.exceljava.com4j.office.MsoTriState relativeToOriginalSize);

  /**
   * @param factor Mandatory float parameter.
   * @param relativeToOriginalSize Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param fScale Optional parameter. Default value is 0
   */

  @DISPID(19) //= 0x13. The runtime will prefer the VTID if present
  @VTID(19)
  void scaleHeight(
    float factor,
    com.exceljava.com4j.office.MsoTriState relativeToOriginalSize,
    @Optional @DefaultValue("0") com.exceljava.com4j.office.MsoScaleFrom fScale);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoScaleFrom parameter fScale is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * scaleWidth(factor, relativeToOriginalSize, 0);
   * </code>
   * </p>
   * @param factor Mandatory float parameter.
   * @param relativeToOriginalSize Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(20) //= 0x14. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {com.exceljava.com4j.office.MsoScaleFrom.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  void scaleWidth(
    float factor,
    com.exceljava.com4j.office.MsoTriState relativeToOriginalSize);

  /**
   * @param factor Mandatory float parameter.
   * @param relativeToOriginalSize Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param fScale Optional parameter. Default value is 0
   */

  @DISPID(20) //= 0x14. The runtime will prefer the VTID if present
  @VTID(20)
  void scaleWidth(
    float factor,
    com.exceljava.com4j.office.MsoTriState relativeToOriginalSize,
    @Optional @DefaultValue("0") com.exceljava.com4j.office.MsoScaleFrom fScale);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter replace is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * select(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(21) //= 0x15. The runtime will prefer the VTID if present
  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void select();

  /**
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(21) //= 0x15. The runtime will prefer the VTID if present
  @VTID(21)
  void select(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object replace);


  /**
   */

  @DISPID(22) //= 0x16. The runtime will prefer the VTID if present
  @VTID(22)
  void setShapesDefaultProperties();


  /**
   * @return  Returns a value of type com.exceljava.com4j.office.ShapeRange
   */

  @DISPID(23) //= 0x17. The runtime will prefer the VTID if present
  @VTID(23)
  com.exceljava.com4j.office.ShapeRange ungroup();


  @VTID(23)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.ShapeRange.class})
  com.exceljava.com4j.office.Shape ungroup(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param zOrderCmd Mandatory com.exceljava.com4j.office.MsoZOrderCmd parameter.
   */

  @DISPID(24) //= 0x18. The runtime will prefer the VTID if present
  @VTID(24)
  void zOrder(
    com.exceljava.com4j.office.MsoZOrderCmd zOrderCmd);


  /**
   * <p>
   * Getter method for the COM property "Adjustments"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Adjustments
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(25)
  com.exceljava.com4j.office.Adjustments getAdjustments();


  @VTID(25)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.Adjustments.class})
  float getAdjustments(
    int index);

  /**
   * <p>
   * Getter method for the COM property "AutoShapeType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoAutoShapeType
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(26)
  com.exceljava.com4j.office.MsoAutoShapeType getAutoShapeType();


  /**
   * <p>
   * Setter method for the COM property "AutoShapeType"
   * </p>
   * @param autoShapeType Mandatory com.exceljava.com4j.office.MsoAutoShapeType parameter.
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(27)
  void setAutoShapeType(
    com.exceljava.com4j.office.MsoAutoShapeType autoShapeType);


  /**
   * <p>
   * Getter method for the COM property "BlackWhiteMode"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoBlackWhiteMode
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(28)
  com.exceljava.com4j.office.MsoBlackWhiteMode getBlackWhiteMode();


  /**
   * <p>
   * Setter method for the COM property "BlackWhiteMode"
   * </p>
   * @param blackWhiteMode Mandatory com.exceljava.com4j.office.MsoBlackWhiteMode parameter.
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(29)
  void setBlackWhiteMode(
    com.exceljava.com4j.office.MsoBlackWhiteMode blackWhiteMode);


  /**
   * <p>
   * Getter method for the COM property "Callout"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CalloutFormat
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(30)
  com.exceljava.com4j.office.CalloutFormat getCallout();


  /**
   * <p>
   * Getter method for the COM property "ConnectionSiteCount"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(104) //= 0x68. The runtime will prefer the VTID if present
  @VTID(31)
  int getConnectionSiteCount();


  /**
   * <p>
   * Getter method for the COM property "Connector"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(32)
  com.exceljava.com4j.office.MsoTriState getConnector();


  /**
   * <p>
   * Getter method for the COM property "ConnectorFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.ConnectorFormat
   */

  @DISPID(106) //= 0x6a. The runtime will prefer the VTID if present
  @VTID(33)
  com.exceljava.com4j.office.ConnectorFormat getConnectorFormat();


  /**
   * <p>
   * Getter method for the COM property "Fill"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.FillFormat
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(34)
  com.exceljava.com4j.office.FillFormat getFill();


  /**
   * <p>
   * Getter method for the COM property "GroupItems"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.GroupShapes
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(35)
  com.exceljava.com4j.office.GroupShapes getGroupItems();


  @VTID(35)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.GroupShapes.class})
  com.exceljava.com4j.office.Shape getGroupItems(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(36)
  float getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param height Mandatory float parameter.
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(37)
  void setHeight(
    float height);


  /**
   * <p>
   * Getter method for the COM property "HorizontalFlip"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(110) //= 0x6e. The runtime will prefer the VTID if present
  @VTID(38)
  com.exceljava.com4j.office.MsoTriState getHorizontalFlip();


  /**
   * <p>
   * Getter method for the COM property "Left"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(111) //= 0x6f. The runtime will prefer the VTID if present
  @VTID(39)
  float getLeft();


  /**
   * <p>
   * Setter method for the COM property "Left"
   * </p>
   * @param left Mandatory float parameter.
   */

  @DISPID(111) //= 0x6f. The runtime will prefer the VTID if present
  @VTID(40)
  void setLeft(
    float left);


  /**
   * <p>
   * Getter method for the COM property "Line"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.LineFormat
   */

  @DISPID(112) //= 0x70. The runtime will prefer the VTID if present
  @VTID(41)
  com.exceljava.com4j.office.LineFormat getLine();


  /**
   * <p>
   * Getter method for the COM property "LockAspectRatio"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(113) //= 0x71. The runtime will prefer the VTID if present
  @VTID(42)
  com.exceljava.com4j.office.MsoTriState getLockAspectRatio();


  /**
   * <p>
   * Setter method for the COM property "LockAspectRatio"
   * </p>
   * @param lockAspectRatio Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(113) //= 0x71. The runtime will prefer the VTID if present
  @VTID(43)
  void setLockAspectRatio(
    com.exceljava.com4j.office.MsoTriState lockAspectRatio);


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(115) //= 0x73. The runtime will prefer the VTID if present
  @VTID(44)
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param name Mandatory java.lang.String parameter.
   */

  @DISPID(115) //= 0x73. The runtime will prefer the VTID if present
  @VTID(45)
  void setName(
    java.lang.String name);


  /**
   * <p>
   * Getter method for the COM property "Nodes"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.ShapeNodes
   */

  @DISPID(116) //= 0x74. The runtime will prefer the VTID if present
  @VTID(46)
  com.exceljava.com4j.office.ShapeNodes getNodes();


  @VTID(46)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.ShapeNodes.class})
  com.exceljava.com4j.excel.ShapeNode getNodes(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "Rotation"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(117) //= 0x75. The runtime will prefer the VTID if present
  @VTID(47)
  float getRotation();


  /**
   * <p>
   * Setter method for the COM property "Rotation"
   * </p>
   * @param rotation Mandatory float parameter.
   */

  @DISPID(117) //= 0x75. The runtime will prefer the VTID if present
  @VTID(48)
  void setRotation(
    float rotation);


  /**
   * <p>
   * Getter method for the COM property "PictureFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.PictureFormat
   */

  @DISPID(118) //= 0x76. The runtime will prefer the VTID if present
  @VTID(49)
  com.exceljava.com4j.office.PictureFormat getPictureFormat();


  /**
   * <p>
   * Getter method for the COM property "Shadow"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.ShadowFormat
   */

  @DISPID(119) //= 0x77. The runtime will prefer the VTID if present
  @VTID(50)
  com.exceljava.com4j.office.ShadowFormat getShadow();


  /**
   * <p>
   * Getter method for the COM property "TextEffect"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.TextEffectFormat
   */

  @DISPID(120) //= 0x78. The runtime will prefer the VTID if present
  @VTID(51)
  com.exceljava.com4j.office.TextEffectFormat getTextEffect();


  /**
   * <p>
   * Getter method for the COM property "TextFrame"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.TextFrame
   */

  @DISPID(121) //= 0x79. The runtime will prefer the VTID if present
  @VTID(52)
  com.exceljava.com4j.office.TextFrame getTextFrame();


  /**
   * <p>
   * Getter method for the COM property "ThreeD"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.ThreeDFormat
   */

  @DISPID(122) //= 0x7a. The runtime will prefer the VTID if present
  @VTID(53)
  com.exceljava.com4j.office.ThreeDFormat getThreeD();


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(123) //= 0x7b. The runtime will prefer the VTID if present
  @VTID(54)
  float getTop();


  /**
   * <p>
   * Setter method for the COM property "Top"
   * </p>
   * @param top Mandatory float parameter.
   */

  @DISPID(123) //= 0x7b. The runtime will prefer the VTID if present
  @VTID(55)
  void setTop(
    float top);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoShapeType
   */

  @DISPID(124) //= 0x7c. The runtime will prefer the VTID if present
  @VTID(56)
  com.exceljava.com4j.office.MsoShapeType getType();


  /**
   * <p>
   * Getter method for the COM property "VerticalFlip"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(125) //= 0x7d. The runtime will prefer the VTID if present
  @VTID(57)
  com.exceljava.com4j.office.MsoTriState getVerticalFlip();


  /**
   * <p>
   * Getter method for the COM property "Vertices"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(126) //= 0x7e. The runtime will prefer the VTID if present
  @VTID(58)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getVertices();


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(127) //= 0x7f. The runtime will prefer the VTID if present
  @VTID(59)
  com.exceljava.com4j.office.MsoTriState getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param visible Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(127) //= 0x7f. The runtime will prefer the VTID if present
  @VTID(60)
  void setVisible(
    com.exceljava.com4j.office.MsoTriState visible);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(128) //= 0x80. The runtime will prefer the VTID if present
  @VTID(61)
  float getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param width Mandatory float parameter.
   */

  @DISPID(128) //= 0x80. The runtime will prefer the VTID if present
  @VTID(62)
  void setWidth(
    float width);


  /**
   * <p>
   * Getter method for the COM property "ZOrderPosition"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(129) //= 0x81. The runtime will prefer the VTID if present
  @VTID(63)
  int getZOrderPosition();


  /**
   * <p>
   * Getter method for the COM property "Script"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Script
   */

  @DISPID(130) //= 0x82. The runtime will prefer the VTID if present
  @VTID(64)
  com.exceljava.com4j.office.Script getScript();


  /**
   * <p>
   * Getter method for the COM property "AlternativeText"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(131) //= 0x83. The runtime will prefer the VTID if present
  @VTID(65)
  java.lang.String getAlternativeText();


  /**
   * <p>
   * Setter method for the COM property "AlternativeText"
   * </p>
   * @param alternativeText Mandatory java.lang.String parameter.
   */

  @DISPID(131) //= 0x83. The runtime will prefer the VTID if present
  @VTID(66)
  void setAlternativeText(
    java.lang.String alternativeText);


  /**
   * <p>
   * Getter method for the COM property "HasDiagram"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(132) //= 0x84. The runtime will prefer the VTID if present
  @VTID(67)
  com.exceljava.com4j.office.MsoTriState getHasDiagram();


  /**
   * <p>
   * Getter method for the COM property "Diagram"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoDiagram
   */

  @DISPID(133) //= 0x85. The runtime will prefer the VTID if present
  @VTID(68)
  com.exceljava.com4j.office.IMsoDiagram getDiagram();


  /**
   * <p>
   * Getter method for the COM property "HasDiagramNode"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(134) //= 0x86. The runtime will prefer the VTID if present
  @VTID(69)
  com.exceljava.com4j.office.MsoTriState getHasDiagramNode();


  /**
   * <p>
   * Getter method for the COM property "DiagramNode"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.DiagramNode
   */

  @DISPID(135) //= 0x87. The runtime will prefer the VTID if present
  @VTID(70)
  com.exceljava.com4j.office.DiagramNode getDiagramNode();


  /**
   * <p>
   * Getter method for the COM property "Child"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(136) //= 0x88. The runtime will prefer the VTID if present
  @VTID(71)
  com.exceljava.com4j.office.MsoTriState getChild();


  /**
   * <p>
   * Getter method for the COM property "ParentGroup"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(137) //= 0x89. The runtime will prefer the VTID if present
  @VTID(72)
  com.exceljava.com4j.office.Shape getParentGroup();


  /**
   * <p>
   * Getter method for the COM property "CanvasItems"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CanvasShapes
   */

  @DISPID(138) //= 0x8a. The runtime will prefer the VTID if present
  @VTID(73)
  com.exceljava.com4j.office.CanvasShapes getCanvasItems();


  @VTID(73)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.CanvasShapes.class})
  com.exceljava.com4j.office.Shape getCanvasItems(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "Id"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(139) //= 0x8b. The runtime will prefer the VTID if present
  @VTID(74)
  int getId();


  /**
   * @param increment Mandatory float parameter.
   */

  @DISPID(140) //= 0x8c. The runtime will prefer the VTID if present
  @VTID(75)
  void canvasCropLeft(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @DISPID(141) //= 0x8d. The runtime will prefer the VTID if present
  @VTID(76)
  void canvasCropTop(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @DISPID(142) //= 0x8e. The runtime will prefer the VTID if present
  @VTID(77)
  void canvasCropRight(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @DISPID(143) //= 0x8f. The runtime will prefer the VTID if present
  @VTID(78)
  void canvasCropBottom(
    float increment);


  /**
   * <p>
   * Setter method for the COM property "RTF"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(144) //= 0x90. The runtime will prefer the VTID if present
  @VTID(79)
  void setRTF(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "TextFrame2"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.TextFrame2
   */

  @DISPID(145) //= 0x91. The runtime will prefer the VTID if present
  @VTID(80)
  com.exceljava.com4j.office.TextFrame2 getTextFrame2();


  /**
   */

  @DISPID(146) //= 0x92. The runtime will prefer the VTID if present
  @VTID(81)
  void cut();


  /**
   */

  @DISPID(147) //= 0x93. The runtime will prefer the VTID if present
  @VTID(82)
  void copy();


  /**
   * <p>
   * Getter method for the COM property "HasChart"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(148) //= 0x94. The runtime will prefer the VTID if present
  @VTID(83)
  com.exceljava.com4j.office.MsoTriState getHasChart();


  /**
   * <p>
   * Getter method for the COM property "Chart"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChart
   */

  @DISPID(149) //= 0x95. The runtime will prefer the VTID if present
  @VTID(84)
  com.exceljava.com4j.office.IMsoChart getChart();


  /**
   * <p>
   * Getter method for the COM property "ShapeStyle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoShapeStyleIndex
   */

  @DISPID(150) //= 0x96. The runtime will prefer the VTID if present
  @VTID(85)
  com.exceljava.com4j.office.MsoShapeStyleIndex getShapeStyle();


  /**
   * <p>
   * Setter method for the COM property "ShapeStyle"
   * </p>
   * @param shapeStyle Mandatory com.exceljava.com4j.office.MsoShapeStyleIndex parameter.
   */

  @DISPID(150) //= 0x96. The runtime will prefer the VTID if present
  @VTID(86)
  void setShapeStyle(
    com.exceljava.com4j.office.MsoShapeStyleIndex shapeStyle);


  /**
   * <p>
   * Getter method for the COM property "BackgroundStyle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoBackgroundStyleIndex
   */

  @DISPID(151) //= 0x97. The runtime will prefer the VTID if present
  @VTID(87)
  com.exceljava.com4j.office.MsoBackgroundStyleIndex getBackgroundStyle();


  /**
   * <p>
   * Setter method for the COM property "BackgroundStyle"
   * </p>
   * @param backgroundStyle Mandatory com.exceljava.com4j.office.MsoBackgroundStyleIndex parameter.
   */

  @DISPID(151) //= 0x97. The runtime will prefer the VTID if present
  @VTID(88)
  void setBackgroundStyle(
    com.exceljava.com4j.office.MsoBackgroundStyleIndex backgroundStyle);


  /**
   * <p>
   * Getter method for the COM property "SoftEdge"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SoftEdgeFormat
   */

  @DISPID(152) //= 0x98. The runtime will prefer the VTID if present
  @VTID(89)
  com.exceljava.com4j.office.SoftEdgeFormat getSoftEdge();


  /**
   * <p>
   * Getter method for the COM property "Glow"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.GlowFormat
   */

  @DISPID(153) //= 0x99. The runtime will prefer the VTID if present
  @VTID(90)
  com.exceljava.com4j.office.GlowFormat getGlow();


  /**
   * <p>
   * Getter method for the COM property "Reflection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.ReflectionFormat
   */

  @DISPID(154) //= 0x9a. The runtime will prefer the VTID if present
  @VTID(91)
  com.exceljava.com4j.office.ReflectionFormat getReflection();


  /**
   * <p>
   * Getter method for the COM property "HasSmartArt"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(155) //= 0x9b. The runtime will prefer the VTID if present
  @VTID(92)
  com.exceljava.com4j.office.MsoTriState getHasSmartArt();


  /**
   * <p>
   * Getter method for the COM property "SmartArt"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SmartArt
   */

  @DISPID(156) //= 0x9c. The runtime will prefer the VTID if present
  @VTID(93)
  com.exceljava.com4j.office.SmartArt getSmartArt();


  /**
   * @param layout Mandatory com.exceljava.com4j.office.SmartArtLayout parameter.
   */

  @DISPID(157) //= 0x9d. The runtime will prefer the VTID if present
  @VTID(94)
  void convertTextToSmartArt(
    com.exceljava.com4j.office.SmartArtLayout layout);


  /**
   * <p>
   * Getter method for the COM property "Title"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(158) //= 0x9e. The runtime will prefer the VTID if present
  @VTID(95)
  java.lang.String getTitle();


  /**
   * <p>
   * Setter method for the COM property "Title"
   * </p>
   * @param title Mandatory java.lang.String parameter.
   */

  @DISPID(158) //= 0x9e. The runtime will prefer the VTID if present
  @VTID(96)
  void setTitle(
    java.lang.String title);


  /**
   * <p>
   * Getter method for the COM property "GraphicStyle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoGraphicStyleIndex
   */

  @DISPID(159) //= 0x9f. The runtime will prefer the VTID if present
  @VTID(97)
  com.exceljava.com4j.office.MsoGraphicStyleIndex getGraphicStyle();


  /**
   * <p>
   * Setter method for the COM property "GraphicStyle"
   * </p>
   * @param graphicStyle Mandatory com.exceljava.com4j.office.MsoGraphicStyleIndex parameter.
   */

  @DISPID(159) //= 0x9f. The runtime will prefer the VTID if present
  @VTID(98)
  void setGraphicStyle(
    com.exceljava.com4j.office.MsoGraphicStyleIndex graphicStyle);


  /**
   * @param pictureType Mandatory com.exceljava.com4j.office.MsoPictureType parameter.
   * @param fileName Mandatory java.lang.String parameter.
   * @param fSaveShapesIndividually Mandatory boolean parameter.
   */

  @DISPID(160) //= 0xa0. The runtime will prefer the VTID if present
  @VTID(99)
  void saveAsPicture(
    com.exceljava.com4j.office.MsoPictureType pictureType,
    java.lang.String fileName,
    boolean fSaveShapesIndividually);


  /**
   * <p>
   * Getter method for the COM property "Model3D"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Model3DFormat
   */

  @DISPID(161) //= 0xa1. The runtime will prefer the VTID if present
  @VTID(100)
  com.exceljava.com4j.office.Model3DFormat getModel3D();


  /**
   * <p>
   * Getter method for the COM property "Decorative"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(162) //= 0xa2. The runtime will prefer the VTID if present
  @VTID(101)
  com.exceljava.com4j.office.MsoTriState getDecorative();


  /**
   * <p>
   * Setter method for the COM property "Decorative"
   * </p>
   * @param fDecorative Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(162) //= 0xa2. The runtime will prefer the VTID if present
  @VTID(102)
  void setDecorative(
    com.exceljava.com4j.office.MsoTriState fDecorative);


  // Properties:
}
