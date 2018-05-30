package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0371-0000-0000-C000-000000000046}")
public interface CanvasShapes extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(10)
  int getCount();


  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(11)
  @DefaultMethod
  com.exceljava.com4j.office.Shape item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(12)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * @param type Mandatory com.exceljava.com4j.office.MsoCalloutType parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @param width Mandatory float parameter.
   * @param height Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(13)
  com.exceljava.com4j.office.Shape addCallout(
    com.exceljava.com4j.office.MsoCalloutType type,
    float left,
    float top,
    float width,
    float height);


  /**
   * @param type Mandatory com.exceljava.com4j.office.MsoConnectorType parameter.
   * @param beginX Mandatory float parameter.
   * @param beginY Mandatory float parameter.
   * @param endX Mandatory float parameter.
   * @param endY Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.office.Shape addConnector(
    com.exceljava.com4j.office.MsoConnectorType type,
    float beginX,
    float beginY,
    float endX,
    float endY);


  /**
   * @param safeArrayOfPoints Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.office.Shape addCurve(
    @MarshalAs(NativeType.VARIANT) java.lang.Object safeArrayOfPoints);


  /**
   * @param orientation Mandatory com.exceljava.com4j.office.MsoTextOrientation parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @param width Mandatory float parameter.
   * @param height Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(16)
  com.exceljava.com4j.office.Shape addLabel(
    com.exceljava.com4j.office.MsoTextOrientation orientation,
    float left,
    float top,
    float width,
    float height);


  /**
   * @param beginX Mandatory float parameter.
   * @param beginY Mandatory float parameter.
   * @param endX Mandatory float parameter.
   * @param endY Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(17)
  com.exceljava.com4j.office.Shape addLine(
    float beginX,
    float beginY,
    float endX,
    float endY);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter width is set to -1.0f</li><li>float parameter height is set to -1.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addPicture(fileName, linkToFile, saveWithDocument, left, top, -1.0f, -1.0f);
   * </code>
   * </p>
   * @param fileName Mandatory java.lang.String parameter.
   * @param linkToFile Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param saveWithDocument Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 7}, optParamIndex = {5, 6}, javaType = {float.class, float.class}, nativeType = {NativeType.Float, NativeType.Float}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_R4}, literal = {"-1.0", "-1.0"})
  @ReturnValue(index=7)
  com.exceljava.com4j.office.Shape addPicture(
    java.lang.String fileName,
    com.exceljava.com4j.office.MsoTriState linkToFile,
    com.exceljava.com4j.office.MsoTriState saveWithDocument,
    float left,
    float top);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter height is set to -1.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addPicture(fileName, linkToFile, saveWithDocument, left, top, width, -1.0f);
   * </code>
   * </p>
   * @param fileName Mandatory java.lang.String parameter.
   * @param linkToFile Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param saveWithDocument Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @param width Optional parameter. Default value is -1.0f
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(18)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 7}, optParamIndex = {6}, javaType = {float.class}, nativeType = {NativeType.Float}, variantType = {Variant.Type.VT_R4}, literal = {"-1.0"})
  @ReturnValue(index=7)
  com.exceljava.com4j.office.Shape addPicture(
    java.lang.String fileName,
    com.exceljava.com4j.office.MsoTriState linkToFile,
    com.exceljava.com4j.office.MsoTriState saveWithDocument,
    float left,
    float top,
    @Optional @DefaultValue("-1") float width);

  /**
   * @param fileName Mandatory java.lang.String parameter.
   * @param linkToFile Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param saveWithDocument Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @param width Optional parameter. Default value is -1.0f
   * @param height Optional parameter. Default value is -1.0f
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(18)
  com.exceljava.com4j.office.Shape addPicture(
    java.lang.String fileName,
    com.exceljava.com4j.office.MsoTriState linkToFile,
    com.exceljava.com4j.office.MsoTriState saveWithDocument,
    float left,
    float top,
    @Optional @DefaultValue("-1") float width,
    @Optional @DefaultValue("-1") float height);


  /**
   * @param safeArrayOfPoints Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(16) //= 0x10. The runtime will prefer the VTID if present
  @VTID(19)
  com.exceljava.com4j.office.Shape addPolyline(
    @MarshalAs(NativeType.VARIANT) java.lang.Object safeArrayOfPoints);


  /**
   * @param type Mandatory com.exceljava.com4j.office.MsoAutoShapeType parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @param width Mandatory float parameter.
   * @param height Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(17) //= 0x11. The runtime will prefer the VTID if present
  @VTID(20)
  com.exceljava.com4j.office.Shape addShape(
    com.exceljava.com4j.office.MsoAutoShapeType type,
    float left,
    float top,
    float width,
    float height);


  /**
   * @param presetTextEffect Mandatory com.exceljava.com4j.office.MsoPresetTextEffect parameter.
   * @param text Mandatory java.lang.String parameter.
   * @param fontName Mandatory java.lang.String parameter.
   * @param fontSize Mandatory float parameter.
   * @param fontBold Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param fontItalic Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(18) //= 0x12. The runtime will prefer the VTID if present
  @VTID(21)
  com.exceljava.com4j.office.Shape addTextEffect(
    com.exceljava.com4j.office.MsoPresetTextEffect presetTextEffect,
    java.lang.String text,
    java.lang.String fontName,
    float fontSize,
    com.exceljava.com4j.office.MsoTriState fontBold,
    com.exceljava.com4j.office.MsoTriState fontItalic,
    float left,
    float top);


  /**
   * @param orientation Mandatory com.exceljava.com4j.office.MsoTextOrientation parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @param width Mandatory float parameter.
   * @param height Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(19) //= 0x13. The runtime will prefer the VTID if present
  @VTID(22)
  com.exceljava.com4j.office.Shape addTextbox(
    com.exceljava.com4j.office.MsoTextOrientation orientation,
    float left,
    float top,
    float width,
    float height);


  /**
   * @param editingType Mandatory com.exceljava.com4j.office.MsoEditingType parameter.
   * @param x1 Mandatory float parameter.
   * @param y1 Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.FreeformBuilder
   */

  @DISPID(20) //= 0x14. The runtime will prefer the VTID if present
  @VTID(23)
  com.exceljava.com4j.office.FreeformBuilder buildFreeform(
    com.exceljava.com4j.office.MsoEditingType editingType,
    float x1,
    float y1);


  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.ShapeRange
   */

  @DISPID(21) //= 0x15. The runtime will prefer the VTID if present
  @VTID(24)
  com.exceljava.com4j.office.ShapeRange range(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   */

  @DISPID(22) //= 0x16. The runtime will prefer the VTID if present
  @VTID(25)
  void selectAll();


  /**
   * <p>
   * Getter method for the COM property "Background"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(26)
  com.exceljava.com4j.office.Shape getBackground();


  // Properties:
}
