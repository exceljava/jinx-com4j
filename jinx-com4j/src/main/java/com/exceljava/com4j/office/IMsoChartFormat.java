package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C1730-0000-0000-C000-000000000046}")
public interface IMsoChartFormat extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Fill"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.FillFormat
   */

  @DISPID(1610743808) //= 0x60020000. The runtime will prefer the VTID if present
  @VTID(7)
  com.exceljava.com4j.office.FillFormat getFill();


  /**
   * <p>
   * Getter method for the COM property "Glow"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.GlowFormat
   */

  @DISPID(1610743809) //= 0x60020001. The runtime will prefer the VTID if present
  @VTID(8)
  com.exceljava.com4j.office.GlowFormat getGlow();


  /**
   * <p>
   * Getter method for the COM property "Line"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.LineFormat
   */

  @DISPID(1610743810) //= 0x60020002. The runtime will prefer the VTID if present
  @VTID(9)
  com.exceljava.com4j.office.LineFormat getLine();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610743811) //= 0x60020003. The runtime will prefer the VTID if present
  @VTID(10)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "PictureFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.PictureFormat
   */

  @DISPID(1610743812) //= 0x60020004. The runtime will prefer the VTID if present
  @VTID(11)
  com.exceljava.com4j.office.PictureFormat getPictureFormat();


  /**
   * <p>
   * Getter method for the COM property "Shadow"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.ShadowFormat
   */

  @DISPID(1610743813) //= 0x60020005. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.ShadowFormat getShadow();


  /**
   * <p>
   * Getter method for the COM property "SoftEdge"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.SoftEdgeFormat
   */

  @DISPID(1610743814) //= 0x60020006. The runtime will prefer the VTID if present
  @VTID(13)
  com.exceljava.com4j.office.SoftEdgeFormat getSoftEdge();


  /**
   * <p>
   * Getter method for the COM property "TextFrame2"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.TextFrame2
   */

  @DISPID(1610743815) //= 0x60020007. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.office.TextFrame2 getTextFrame2();


  /**
   * <p>
   * Getter method for the COM property "ThreeD"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.ThreeDFormat
   */

  @DISPID(1610743816) //= 0x60020008. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.office.ThreeDFormat getThreeD();


  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(148) //= 0x94. The runtime will prefer the VTID if present
  @VTID(16)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(149) //= 0x95. The runtime will prefer the VTID if present
  @VTID(17)
  int getCreator();


  /**
   * <p>
   * Getter method for the COM property "Adjustments"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Adjustments
   */

  @DISPID(200) //= 0xc8. The runtime will prefer the VTID if present
  @VTID(18)
  com.exceljava.com4j.office.Adjustments getAdjustments();


  @VTID(18)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.Adjustments.class})
  float getAdjustments(
    int index);

  /**
   * <p>
   * Getter method for the COM property "AutoShapeType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoAutoShapeType
   */

  @DISPID(201) //= 0xc9. The runtime will prefer the VTID if present
  @VTID(19)
  com.exceljava.com4j.office.MsoAutoShapeType getAutoShapeType();


  /**
   * <p>
   * Setter method for the COM property "AutoShapeType"
   * </p>
   * @param autoShapeType Mandatory com.exceljava.com4j.office.MsoAutoShapeType parameter.
   */

  @DISPID(201) //= 0xc9. The runtime will prefer the VTID if present
  @VTID(20)
  void setAutoShapeType(
    com.exceljava.com4j.office.MsoAutoShapeType autoShapeType);


  // Properties:
}
