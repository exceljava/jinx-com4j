package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024439-0001-0000-C000-000000000046}")
public interface IShape extends Com4jObject {
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
   */

  @VTID(10)
  void apply();


  /**
   */

  @VTID(11)
  void delete();


  /**
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(12)
  com.exceljava.com4j.excel.Shape duplicate();


  /**
   * @param flipCmd Mandatory com.exceljava.com4j.office.MsoFlipCmd parameter.
   */

  @VTID(13)
  void flip(
    com.exceljava.com4j.office.MsoFlipCmd flipCmd);


  /**
   * @param increment Mandatory float parameter.
   */

  @VTID(14)
  void incrementLeft(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @VTID(15)
  void incrementRotation(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @VTID(16)
  void incrementTop(
    float increment);


  /**
   */

  @VTID(17)
  void pickUp();


  /**
   */

  @VTID(18)
  void rerouteConnections();


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

  @VTID(19)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void scaleHeight(
    float factor,
    com.exceljava.com4j.office.MsoTriState relativeToOriginalSize);

  /**
   * @param factor Mandatory float parameter.
   * @param relativeToOriginalSize Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param scale Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(19)
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

  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void scaleWidth(
    float factor,
    com.exceljava.com4j.office.MsoTriState relativeToOriginalSize);

  /**
   * @param factor Mandatory float parameter.
   * @param relativeToOriginalSize Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param scale Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(20)
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

  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void select();

  /**
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(21)
  void select(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object replace);


  /**
   */

  @VTID(22)
  void setShapesDefaultProperties();


  /**
   * @return  Returns a value of type com.exceljava.com4j.excel.ShapeRange
   */

  @VTID(23)
  com.exceljava.com4j.excel.ShapeRange ungroup();


  /**
   * @param zOrderCmd Mandatory com.exceljava.com4j.office.MsoZOrderCmd parameter.
   */

  @VTID(24)
  void zOrder(
    com.exceljava.com4j.office.MsoZOrderCmd zOrderCmd);


  /**
   * <p>
   * Getter method for the COM property "Adjustments"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Adjustments
   */

  @VTID(25)
  com.exceljava.com4j.excel.Adjustments getAdjustments();


  @VTID(25)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.excel.Adjustments.class})
  float getAdjustments(
    int index);

  /**
   * <p>
   * Getter method for the COM property "TextFrame"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.TextFrame
   */

  @VTID(26)
  com.exceljava.com4j.excel.TextFrame getTextFrame();


  /**
   * <p>
   * Getter method for the COM property "AutoShapeType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoAutoShapeType
   */

  @VTID(27)
  com.exceljava.com4j.office.MsoAutoShapeType getAutoShapeType();


  /**
   * <p>
   * Setter method for the COM property "AutoShapeType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoAutoShapeType parameter.
   */

  @VTID(28)
  void setAutoShapeType(
    com.exceljava.com4j.office.MsoAutoShapeType rhs);


  /**
   * <p>
   * Getter method for the COM property "Callout"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.CalloutFormat
   */

  @VTID(29)
  com.exceljava.com4j.excel.CalloutFormat getCallout();


  /**
   * <p>
   * Getter method for the COM property "ConnectionSiteCount"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(30)
  int getConnectionSiteCount();


  /**
   * <p>
   * Getter method for the COM property "Connector"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(31)
  com.exceljava.com4j.office.MsoTriState getConnector();


  /**
   * <p>
   * Getter method for the COM property "ConnectorFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ConnectorFormat
   */

  @VTID(32)
  com.exceljava.com4j.excel.ConnectorFormat getConnectorFormat();


  /**
   * <p>
   * Getter method for the COM property "Fill"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.FillFormat
   */

  @VTID(33)
  com.exceljava.com4j.excel.FillFormat getFill();


  /**
   * <p>
   * Getter method for the COM property "GroupItems"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.GroupShapes
   */

  @VTID(34)
  com.exceljava.com4j.excel.GroupShapes getGroupItems();


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(35)
  float getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(36)
  void setHeight(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "HorizontalFlip"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(37)
  com.exceljava.com4j.office.MsoTriState getHorizontalFlip();


  /**
   * <p>
   * Getter method for the COM property "Left"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(38)
  float getLeft();


  /**
   * <p>
   * Setter method for the COM property "Left"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(39)
  void setLeft(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "Line"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.LineFormat
   */

  @VTID(40)
  com.exceljava.com4j.excel.LineFormat getLine();


  /**
   * <p>
   * Getter method for the COM property "LockAspectRatio"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(41)
  com.exceljava.com4j.office.MsoTriState getLockAspectRatio();


  /**
   * <p>
   * Setter method for the COM property "LockAspectRatio"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @VTID(42)
  void setLockAspectRatio(
    com.exceljava.com4j.office.MsoTriState rhs);


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(43)
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(44)
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Nodes"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ShapeNodes
   */

  @VTID(45)
  com.exceljava.com4j.excel.ShapeNodes getNodes();


  @VTID(45)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.excel.ShapeNodes.class})
  com.exceljava.com4j.excel.ShapeNode getNodes(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "Rotation"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(46)
  float getRotation();


  /**
   * <p>
   * Setter method for the COM property "Rotation"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(47)
  void setRotation(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "PictureFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PictureFormat
   */

  @VTID(48)
  com.exceljava.com4j.excel.PictureFormat getPictureFormat();


  /**
   * <p>
   * Getter method for the COM property "Shadow"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ShadowFormat
   */

  @VTID(49)
  com.exceljava.com4j.excel.ShadowFormat getShadow();


  /**
   * <p>
   * Getter method for the COM property "TextEffect"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.TextEffectFormat
   */

  @VTID(50)
  com.exceljava.com4j.excel.TextEffectFormat getTextEffect();


  /**
   * <p>
   * Getter method for the COM property "ThreeD"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ThreeDFormat
   */

  @VTID(51)
  com.exceljava.com4j.excel.ThreeDFormat getThreeD();


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(52)
  float getTop();


  /**
   * <p>
   * Setter method for the COM property "Top"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(53)
  void setTop(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoShapeType
   */

  @VTID(54)
  com.exceljava.com4j.office.MsoShapeType getType();


  /**
   * <p>
   * Getter method for the COM property "VerticalFlip"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(55)
  com.exceljava.com4j.office.MsoTriState getVerticalFlip();


  /**
   * <p>
   * Getter method for the COM property "Vertices"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(56)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getVertices();


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(57)
  com.exceljava.com4j.office.MsoTriState getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @VTID(58)
  void setVisible(
    com.exceljava.com4j.office.MsoTriState rhs);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(59)
  float getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param rhs Mandatory float parameter.
   */

  @VTID(60)
  void setWidth(
    float rhs);


  /**
   * <p>
   * Getter method for the COM property "ZOrderPosition"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(61)
  int getZOrderPosition();


  /**
   * <p>
   * Getter method for the COM property "Hyperlink"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Hyperlink
   */

  @VTID(62)
  com.exceljava.com4j.excel.Hyperlink getHyperlink();


  /**
   * <p>
   * Getter method for the COM property "BlackWhiteMode"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoBlackWhiteMode
   */

  @VTID(63)
  com.exceljava.com4j.office.MsoBlackWhiteMode getBlackWhiteMode();


  /**
   * <p>
   * Setter method for the COM property "BlackWhiteMode"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoBlackWhiteMode parameter.
   */

  @VTID(64)
  void setBlackWhiteMode(
    com.exceljava.com4j.office.MsoBlackWhiteMode rhs);


  /**
   * <p>
   * Getter method for the COM property "DrawingObject"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(65)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getDrawingObject();


  /**
   * <p>
   * Getter method for the COM property "OnAction"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(66)
  java.lang.String getOnAction();


  /**
   * <p>
   * Setter method for the COM property "OnAction"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(67)
  void setOnAction(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Locked"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(68)
  boolean getLocked();


  /**
   * <p>
   * Setter method for the COM property "Locked"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(69)
  void setLocked(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "TopLeftCell"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(70)
  com.exceljava.com4j.excel.Range getTopLeftCell();


  /**
   * <p>
   * Getter method for the COM property "BottomRightCell"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(71)
  com.exceljava.com4j.excel.Range getBottomRightCell();


  /**
   * <p>
   * Getter method for the COM property "Placement"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlPlacement
   */

  @VTID(72)
  com.exceljava.com4j.excel.XlPlacement getPlacement();


  /**
   * <p>
   * Setter method for the COM property "Placement"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlPlacement parameter.
   */

  @VTID(73)
  void setPlacement(
    com.exceljava.com4j.excel.XlPlacement rhs);


  /**
   */

  @VTID(74)
  void copy();


  /**
   */

  @VTID(75)
  void cut();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter appearance is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copyPicture(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @VTID(76)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void copyPicture();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copyPicture(appearance, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param appearance Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(76)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void copyPicture(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object appearance);

  /**
   * @param appearance Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(76)
  void copyPicture(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object appearance,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format);


  /**
   * <p>
   * Getter method for the COM property "ControlFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ControlFormat
   */

  @VTID(77)
  com.exceljava.com4j.excel.ControlFormat getControlFormat();


  /**
   * <p>
   * Getter method for the COM property "LinkFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.LinkFormat
   */

  @VTID(78)
  com.exceljava.com4j.excel.LinkFormat getLinkFormat();


  /**
   * <p>
   * Getter method for the COM property "OLEFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.OLEFormat
   */

  @VTID(79)
  com.exceljava.com4j.excel.OLEFormat getOLEFormat();


  /**
   * <p>
   * Getter method for the COM property "FormControlType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlFormControl
   */

  @VTID(80)
  com.exceljava.com4j.excel.XlFormControl getFormControlType();


  /**
   * <p>
   * Getter method for the COM property "AlternativeText"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(81)
  java.lang.String getAlternativeText();


  /**
   * <p>
   * Setter method for the COM property "AlternativeText"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(82)
  void setAlternativeText(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Script"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Script
   */

  @VTID(83)
  com.exceljava.com4j.office.Script getScript();


  /**
   * <p>
   * Getter method for the COM property "DiagramNode"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.DiagramNode
   */

  @VTID(84)
  com.exceljava.com4j.excel.DiagramNode getDiagramNode();


  /**
   * <p>
   * Getter method for the COM property "HasDiagramNode"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(85)
  com.exceljava.com4j.office.MsoTriState getHasDiagramNode();


  /**
   * <p>
   * Getter method for the COM property "Diagram"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Diagram
   */

  @VTID(86)
  com.exceljava.com4j.excel.Diagram getDiagram();


  /**
   * <p>
   * Getter method for the COM property "HasDiagram"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(87)
  com.exceljava.com4j.office.MsoTriState getHasDiagram();


  /**
   * <p>
   * Getter method for the COM property "Child"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(88)
  com.exceljava.com4j.office.MsoTriState getChild();


  /**
   * <p>
   * Getter method for the COM property "ParentGroup"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(89)
  com.exceljava.com4j.excel.Shape getParentGroup();


  /**
   * <p>
   * Getter method for the COM property "CanvasItems"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CanvasShapes
   */

  @VTID(90)
  com.exceljava.com4j.office.CanvasShapes getCanvasItems();


  @VTID(90)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.CanvasShapes.class})
  com.exceljava.com4j.office.Shape getCanvasItems(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "ID"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(91)
  int getID();


  /**
   * @param increment Mandatory float parameter.
   */

  @VTID(92)
  void canvasCropLeft(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @VTID(93)
  void canvasCropTop(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @VTID(94)
  void canvasCropRight(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @VTID(95)
  void canvasCropBottom(
    float increment);


  /**
   * <p>
   * Getter method for the COM property "Chart"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Chart
   */

  @VTID(96)
  com.exceljava.com4j.excel._Chart getChart();


  /**
   * <p>
   * Getter method for the COM property "HasChart"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(97)
  com.exceljava.com4j.office.MsoTriState getHasChart();


  /**
   * <p>
   * Getter method for the COM property "TextFrame2"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.TextFrame2
   */

  @VTID(98)
  com.exceljava.com4j.excel.TextFrame2 getTextFrame2();


  /**
   * <p>
   * Getter method for the COM property "ShapeStyle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoShapeStyleIndex
   */

  @VTID(99)
  com.exceljava.com4j.office.MsoShapeStyleIndex getShapeStyle();


  /**
   * <p>
   * Setter method for the COM property "ShapeStyle"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoShapeStyleIndex parameter.
   */

  @VTID(100)
  void setShapeStyle(
    com.exceljava.com4j.office.MsoShapeStyleIndex rhs);


  /**
   * <p>
   * Getter method for the COM property "BackgroundStyle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoBackgroundStyleIndex
   */

  @VTID(101)
  com.exceljava.com4j.office.MsoBackgroundStyleIndex getBackgroundStyle();


  /**
   * <p>
   * Setter method for the COM property "BackgroundStyle"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoBackgroundStyleIndex parameter.
   */

  @VTID(102)
  void setBackgroundStyle(
    com.exceljava.com4j.office.MsoBackgroundStyleIndex rhs);


  /**
   * <p>
   * Getter method for the COM property "SoftEdge"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SoftEdgeFormat
   */

  @VTID(103)
  com.exceljava.com4j.office.SoftEdgeFormat getSoftEdge();


  /**
   * <p>
   * Getter method for the COM property "Glow"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.GlowFormat
   */

  @VTID(104)
  com.exceljava.com4j.office.GlowFormat getGlow();


  /**
   * <p>
   * Getter method for the COM property "Reflection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.ReflectionFormat
   */

  @VTID(105)
  com.exceljava.com4j.office.ReflectionFormat getReflection();


  /**
   * <p>
   * Getter method for the COM property "HasSmartArt"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(106)
  com.exceljava.com4j.office.MsoTriState getHasSmartArt();


  /**
   * <p>
   * Getter method for the COM property "SmartArt"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SmartArt
   */

  @VTID(107)
  com.exceljava.com4j.office.SmartArt getSmartArt();


  /**
   * <p>
   * Getter method for the COM property "Title"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(108)
  java.lang.String getTitle();


  /**
   * <p>
   * Setter method for the COM property "Title"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(109)
  void setTitle(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "GraphicStyle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoGraphicStyleIndex
   */

  @VTID(110)
  com.exceljava.com4j.office.MsoGraphicStyleIndex getGraphicStyle();


  /**
   * <p>
   * Setter method for the COM property "GraphicStyle"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoGraphicStyleIndex parameter.
   */

  @VTID(111)
  void setGraphicStyle(
    com.exceljava.com4j.office.MsoGraphicStyleIndex rhs);


  /**
   * <p>
   * Getter method for the COM property "Model3D"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Model3DFormat
   */

  @VTID(112)
  com.exceljava.com4j.excel.Model3DFormat getModel3D();


  /**
   * <p>
   * Getter method for the COM property "Decorative"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(113)
  com.exceljava.com4j.office.MsoTriState getDecorative();


  /**
   * <p>
   * Setter method for the COM property "Decorative"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @VTID(114)
  void setDecorative(
    com.exceljava.com4j.office.MsoTriState rhs);


  // Properties:
}
