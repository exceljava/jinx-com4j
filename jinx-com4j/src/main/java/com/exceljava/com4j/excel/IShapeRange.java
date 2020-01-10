package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002443B-0001-0000-C000-000000000046}")
public interface IShapeRange extends Com4jObject,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(10)
  int getCount();


  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(11)
  com.exceljava.com4j.excel.Shape item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(12)
  @DefaultMethod
  com.exceljava.com4j.excel.Shape _Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(13)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * @param alignCmd Mandatory com.exceljava.com4j.office.MsoAlignCmd parameter.
   * @param relativeTo Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @VTID(14)
  void align(
    com.exceljava.com4j.office.MsoAlignCmd alignCmd,
    com.exceljava.com4j.office.MsoTriState relativeTo);


  /**
   */

  @VTID(15)
  void apply();


  /**
   */

  @VTID(16)
  void delete();


  /**
   * @param distributeCmd Mandatory com.exceljava.com4j.office.MsoDistributeCmd parameter.
   * @param relativeTo Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @VTID(17)
  void distribute(
    com.exceljava.com4j.office.MsoDistributeCmd distributeCmd,
    com.exceljava.com4j.office.MsoTriState relativeTo);


  /**
   * @return  Returns a value of type com.exceljava.com4j.excel.ShapeRange
   */

  @VTID(18)
  com.exceljava.com4j.excel.ShapeRange duplicate();


  /**
   * @param flipCmd Mandatory com.exceljava.com4j.office.MsoFlipCmd parameter.
   */

  @VTID(19)
  void flip(
    com.exceljava.com4j.office.MsoFlipCmd flipCmd);


  /**
   * @param increment Mandatory float parameter.
   */

  @VTID(20)
  void incrementLeft(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @VTID(21)
  void incrementRotation(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @VTID(22)
  void incrementTop(
    float increment);


  /**
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(23)
  com.exceljava.com4j.excel.Shape group();


  /**
   */

  @VTID(24)
  void pickUp();


  /**
   */

  @VTID(25)
  void rerouteConnections();


  /**
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(26)
  com.exceljava.com4j.excel.Shape regroup();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter scale is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * scaleHeight(factor, relativeToOriginalSize, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param factor Mandatory float parameter.
   * @param relativeToOriginalSize Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @VTID(27)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void scaleHeight(
    float factor,
    com.exceljava.com4j.office.MsoTriState relativeToOriginalSize);

  /**
   * @param factor Mandatory float parameter.
   * @param relativeToOriginalSize Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param scale Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(27)
  void scaleHeight(
    float factor,
    com.exceljava.com4j.office.MsoTriState relativeToOriginalSize,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scale);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter scale is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * scaleWidth(factor, relativeToOriginalSize, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param factor Mandatory float parameter.
   * @param relativeToOriginalSize Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @VTID(28)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void scaleWidth(
    float factor,
    com.exceljava.com4j.office.MsoTriState relativeToOriginalSize);

  /**
   * @param factor Mandatory float parameter.
   * @param relativeToOriginalSize Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param scale Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(28)
  void scaleWidth(
    float factor,
    com.exceljava.com4j.office.MsoTriState relativeToOriginalSize,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scale);


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

  @VTID(29)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void select();

  /**
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(29)
  void select(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object replace);


  /**
   */

  @VTID(30)
  void setShapesDefaultProperties();


  /**
   * @return  Returns a value of type com.exceljava.com4j.excel.ShapeRange
   */

  @VTID(31)
  com.exceljava.com4j.excel.ShapeRange ungroup();


  /**
   * @param zOrderCmd Mandatory com.exceljava.com4j.office.MsoZOrderCmd parameter.
   */

  @VTID(32)
  void zOrder(
    com.exceljava.com4j.office.MsoZOrderCmd zOrderCmd);


  /**
   * <p>
   * Getter method for the COM property "Adjustments"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Adjustments
   */

  @VTID(33)
  com.exceljava.com4j.excel.Adjustments getAdjustments();


  @VTID(33)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.excel.Adjustments.class})
  float getAdjustments(
    int index);

  /**
   * <p>
   * Getter method for the COM property "TextFrame"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.TextFrame
   */

  @VTID(34)
  com.exceljava.com4j.excel.TextFrame getTextFrame();


  /**
   * <p>
   * Getter method for the COM property "AutoShapeType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoAutoShapeType
   */

  @VTID(35)
  com.exceljava.com4j.office.MsoAutoShapeType getAutoShapeType();


  /**
   * <p>
   * Setter method for the COM property "AutoShapeType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoAutoShapeType parameter.
   */

  @VTID(36)
  void setAutoShapeType(
    com.exceljava.com4j.office.MsoAutoShapeType rhs);


  /**
   * <p>
   * Getter method for the COM property "Callout"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.CalloutFormat
   */

  @VTID(37)
  com.exceljava.com4j.excel.CalloutFormat getCallout();


  /**
   * <p>
   * Getter method for the COM property "ConnectionSiteCount"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(38)
  int getConnectionSiteCount();


  /**
   * <p>
   * Getter method for the COM property "Connector"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(39)
  com.exceljava.com4j.office.MsoTriState getConnector();


  /**
   * <p>
   * Getter method for the COM property "ConnectorFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ConnectorFormat
   */

  @VTID(40)
  com.exceljava.com4j.excel.ConnectorFormat getConnectorFormat();


  /**
   * <p>
   * Getter method for the COM property "Fill"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.FillFormat
   */

  @VTID(41)
  com.exceljava.com4j.excel.FillFormat getFill();


  /**
   * <p>
   * Getter method for the COM property "GroupItems"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.GroupShapes
   */

  @VTID(42)
  com.exceljava.com4j.excel.GroupShapes getGroupItems();


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(43)
  float getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(44)
  void setHeight(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "HorizontalFlip"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(45)
  com.exceljava.com4j.office.MsoTriState getHorizontalFlip();


  /**
   * <p>
   * Getter method for the COM property "Left"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(46)
  float getLeft();


  /**
   * <p>
   * Setter method for the COM property "Left"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(47)
  void setLeft(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "Line"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.LineFormat
   */

  @VTID(48)
  com.exceljava.com4j.excel.LineFormat getLine();


  /**
   * <p>
   * Getter method for the COM property "LockAspectRatio"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(49)
  com.exceljava.com4j.office.MsoTriState getLockAspectRatio();


  /**
   * <p>
   * Setter method for the COM property "LockAspectRatio"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @VTID(50)
  void setLockAspectRatio(
    com.exceljava.com4j.office.MsoTriState rhs);


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(51)
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(52)
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Nodes"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ShapeNodes
   */

  @VTID(53)
  com.exceljava.com4j.excel.ShapeNodes getNodes();


  @VTID(53)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.excel.ShapeNodes.class})
  com.exceljava.com4j.excel.ShapeNode getNodes(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "Rotation"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(54)
  float getRotation();


  /**
   * <p>
   * Setter method for the COM property "Rotation"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(55)
  void setRotation(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "PictureFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PictureFormat
   */

  @VTID(56)
  com.exceljava.com4j.excel.PictureFormat getPictureFormat();


  /**
   * <p>
   * Getter method for the COM property "Shadow"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ShadowFormat
   */

  @VTID(57)
  com.exceljava.com4j.excel.ShadowFormat getShadow();


  /**
   * <p>
   * Getter method for the COM property "TextEffect"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.TextEffectFormat
   */

  @VTID(58)
  com.exceljava.com4j.excel.TextEffectFormat getTextEffect();


  /**
   * <p>
   * Getter method for the COM property "ThreeD"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ThreeDFormat
   */

  @VTID(59)
  com.exceljava.com4j.excel.ThreeDFormat getThreeD();


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(60)
  float getTop();


  /**
   * <p>
   * Setter method for the COM property "Top"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(61)
  void setTop(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoShapeType
   */

  @VTID(62)
  com.exceljava.com4j.office.MsoShapeType getType();


  /**
   * <p>
   * Getter method for the COM property "VerticalFlip"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(63)
  com.exceljava.com4j.office.MsoTriState getVerticalFlip();


  /**
   * <p>
   * Getter method for the COM property "Vertices"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(64)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getVertices();


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(65)
  com.exceljava.com4j.office.MsoTriState getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @VTID(66)
  void setVisible(
    com.exceljava.com4j.office.MsoTriState rhs);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(67)
  float getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(68)
  void setWidth(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "ZOrderPosition"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(69)
  int getZOrderPosition();


  /**
   * <p>
   * Getter method for the COM property "BlackWhiteMode"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoBlackWhiteMode
   */

  @VTID(70)
  com.exceljava.com4j.office.MsoBlackWhiteMode getBlackWhiteMode();


  /**
   * <p>
   * Setter method for the COM property "BlackWhiteMode"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoBlackWhiteMode parameter.
   */

  @VTID(71)
  void setBlackWhiteMode(
    com.exceljava.com4j.office.MsoBlackWhiteMode rhs);


  /**
   * <p>
   * Getter method for the COM property "AlternativeText"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(72)
  java.lang.String getAlternativeText();


  /**
   * <p>
   * Setter method for the COM property "AlternativeText"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(73)
  void setAlternativeText(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "DiagramNode"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.DiagramNode
   */

  @VTID(74)
  com.exceljava.com4j.excel.DiagramNode getDiagramNode();


  /**
   * <p>
   * Getter method for the COM property "HasDiagramNode"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(75)
  com.exceljava.com4j.office.MsoTriState getHasDiagramNode();


  /**
   * <p>
   * Getter method for the COM property "Diagram"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Diagram
   */

  @VTID(76)
  com.exceljava.com4j.excel.Diagram getDiagram();


  /**
   * <p>
   * Getter method for the COM property "HasDiagram"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(77)
  com.exceljava.com4j.office.MsoTriState getHasDiagram();


  /**
   * <p>
   * Getter method for the COM property "Child"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(78)
  com.exceljava.com4j.office.MsoTriState getChild();


  /**
   * <p>
   * Getter method for the COM property "ParentGroup"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(79)
  com.exceljava.com4j.excel.Shape getParentGroup();


  /**
   * <p>
   * Getter method for the COM property "CanvasItems"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CanvasShapes
   */

  @VTID(80)
  com.exceljava.com4j.office.CanvasShapes getCanvasItems();


  @VTID(80)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.CanvasShapes.class})
  com.exceljava.com4j.office.Shape getCanvasItems(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "ID"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(81)
  int getID();


  /**
   * @param increment Mandatory float parameter.
   */

  @VTID(82)
  void canvasCropLeft(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @VTID(83)
  void canvasCropTop(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @VTID(84)
  void canvasCropRight(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @VTID(85)
  void canvasCropBottom(
    float increment);


  /**
   * <p>
   * Getter method for the COM property "Chart"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Chart
   */

  @VTID(86)
  com.exceljava.com4j.excel._Chart getChart();


  /**
   * <p>
   * Getter method for the COM property "HasChart"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(87)
  com.exceljava.com4j.office.MsoTriState getHasChart();


  /**
   * <p>
   * Getter method for the COM property "TextFrame2"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.TextFrame2
   */

  @VTID(88)
  com.exceljava.com4j.excel.TextFrame2 getTextFrame2();


  /**
   * <p>
   * Getter method for the COM property "ShapeStyle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoShapeStyleIndex
   */

  @VTID(89)
  com.exceljava.com4j.office.MsoShapeStyleIndex getShapeStyle();


  /**
   * <p>
   * Setter method for the COM property "ShapeStyle"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoShapeStyleIndex parameter.
   */

  @VTID(90)
  void setShapeStyle(
    com.exceljava.com4j.office.MsoShapeStyleIndex rhs);


  /**
   * <p>
   * Getter method for the COM property "BackgroundStyle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoBackgroundStyleIndex
   */

  @VTID(91)
  com.exceljava.com4j.office.MsoBackgroundStyleIndex getBackgroundStyle();


  /**
   * <p>
   * Setter method for the COM property "BackgroundStyle"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoBackgroundStyleIndex parameter.
   */

  @VTID(92)
  void setBackgroundStyle(
    com.exceljava.com4j.office.MsoBackgroundStyleIndex rhs);


  /**
   * <p>
   * Getter method for the COM property "SoftEdge"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SoftEdgeFormat
   */

  @VTID(93)
  com.exceljava.com4j.office.SoftEdgeFormat getSoftEdge();


  /**
   * <p>
   * Getter method for the COM property "Glow"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.GlowFormat
   */

  @VTID(94)
  com.exceljava.com4j.office.GlowFormat getGlow();


  /**
   * <p>
   * Getter method for the COM property "Reflection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.ReflectionFormat
   */

  @VTID(95)
  com.exceljava.com4j.office.ReflectionFormat getReflection();


  /**
   * <p>
   * Getter method for the COM property "Title"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(96)
  java.lang.String getTitle();


  /**
   * <p>
   * Setter method for the COM property "Title"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(97)
  void setTitle(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "GraphicStyle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoGraphicStyleIndex
   */

  @VTID(98)
  com.exceljava.com4j.office.MsoGraphicStyleIndex getGraphicStyle();


  /**
   * <p>
   * Setter method for the COM property "GraphicStyle"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoGraphicStyleIndex parameter.
   */

  @VTID(99)
  void setGraphicStyle(
    com.exceljava.com4j.office.MsoGraphicStyleIndex rhs);


  /**
   * <p>
   * Getter method for the COM property "Model3D"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Model3DFormat
   */

  @VTID(100)
  com.exceljava.com4j.excel.Model3DFormat getModel3D();


  /**
   * <p>
   * Getter method for the COM property "Decorative"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(101)
  com.exceljava.com4j.office.MsoTriState getDecorative();


  /**
   * <p>
   * Setter method for the COM property "Decorative"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @VTID(102)
  void setDecorative(
    com.exceljava.com4j.office.MsoTriState rhs);


  // Properties:
}
