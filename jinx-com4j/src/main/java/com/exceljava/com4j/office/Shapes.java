package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C031E-0000-0000-C000-000000000046}")
public interface Shapes extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
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


  /**
   * <p>
   * Getter method for the COM property "Default"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(27)
  com.exceljava.com4j.office.Shape getDefault();


  /**
   * @param type Mandatory com.exceljava.com4j.office.MsoDiagramType parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @param width Mandatory float parameter.
   * @param height Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(23) //= 0x17. The runtime will prefer the VTID if present
  @VTID(28)
  com.exceljava.com4j.office.Shape addDiagram(
    com.exceljava.com4j.office.MsoDiagramType type,
    float left,
    float top,
    float width,
    float height);


  /**
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @param width Mandatory float parameter.
   * @param height Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(25) //= 0x19. The runtime will prefer the VTID if present
  @VTID(29)
  com.exceljava.com4j.office.Shape addCanvas(
    float left,
    float top,
    float width,
    float height);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.XlChartType parameter type is set to -1</li><li>float parameter left is set to -1.0f</li><li>float parameter top is set to -1.0f</li><li>float parameter width is set to -1.0f</li><li>float parameter height is set to -1.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart(-1, -1.0f, -1.0f, -1.0f, -1.0f);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(26) //= 0x1a. The runtime will prefer the VTID if present
  @VTID(30)
  @UseDefaultValues(paramIndexMapping = {5}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {com.exceljava.com4j.office.XlChartType.class, float.class, float.class, float.class, float.class}, nativeType = {NativeType.Int32, NativeType.Float, NativeType.Float, NativeType.Float, NativeType.Float}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4}, literal = {"-1", "-1.0", "-1.0", "-1.0", "-1.0"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.Shape addChart();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter left is set to -1.0f</li><li>float parameter top is set to -1.0f</li><li>float parameter width is set to -1.0f</li><li>float parameter height is set to -1.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart(type, -1.0f, -1.0f, -1.0f, -1.0f);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(26) //= 0x1a. The runtime will prefer the VTID if present
  @VTID(30)
  @UseDefaultValues(paramIndexMapping = {0, 5}, optParamIndex = {1, 2, 3, 4}, javaType = {float.class, float.class, float.class, float.class}, nativeType = {NativeType.Float, NativeType.Float, NativeType.Float, NativeType.Float}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4}, literal = {"-1.0", "-1.0", "-1.0", "-1.0"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.Shape addChart(
    @Optional @DefaultValue("-1") com.exceljava.com4j.office.XlChartType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter top is set to -1.0f</li><li>float parameter width is set to -1.0f</li><li>float parameter height is set to -1.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart(type, left, -1.0f, -1.0f, -1.0f);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is -1
   * @param left Optional parameter. Default value is -1.0f
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(26) //= 0x1a. The runtime will prefer the VTID if present
  @VTID(30)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {float.class, float.class, float.class}, nativeType = {NativeType.Float, NativeType.Float, NativeType.Float}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4}, literal = {"-1.0", "-1.0", "-1.0"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.Shape addChart(
    @Optional @DefaultValue("-1") com.exceljava.com4j.office.XlChartType type,
    @Optional @DefaultValue("-1") float left);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter width is set to -1.0f</li><li>float parameter height is set to -1.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart(type, left, top, -1.0f, -1.0f);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is -1
   * @param left Optional parameter. Default value is -1.0f
   * @param top Optional parameter. Default value is -1.0f
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(26) //= 0x1a. The runtime will prefer the VTID if present
  @VTID(30)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {float.class, float.class}, nativeType = {NativeType.Float, NativeType.Float}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_R4}, literal = {"-1.0", "-1.0"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.Shape addChart(
    @Optional @DefaultValue("-1") com.exceljava.com4j.office.XlChartType type,
    @Optional @DefaultValue("-1") float left,
    @Optional @DefaultValue("-1") float top);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter height is set to -1.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart(type, left, top, width, -1.0f);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is -1
   * @param left Optional parameter. Default value is -1.0f
   * @param top Optional parameter. Default value is -1.0f
   * @param width Optional parameter. Default value is -1.0f
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(26) //= 0x1a. The runtime will prefer the VTID if present
  @VTID(30)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {float.class}, nativeType = {NativeType.Float}, variantType = {Variant.Type.VT_R4}, literal = {"-1.0"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.Shape addChart(
    @Optional @DefaultValue("-1") com.exceljava.com4j.office.XlChartType type,
    @Optional @DefaultValue("-1") float left,
    @Optional @DefaultValue("-1") float top,
    @Optional @DefaultValue("-1") float width);

  /**
   * @param type Optional parameter. Default value is -1
   * @param left Optional parameter. Default value is -1.0f
   * @param top Optional parameter. Default value is -1.0f
   * @param width Optional parameter. Default value is -1.0f
   * @param height Optional parameter. Default value is -1.0f
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(26) //= 0x1a. The runtime will prefer the VTID if present
  @VTID(30)
  com.exceljava.com4j.office.Shape addChart(
    @Optional @DefaultValue("-1") com.exceljava.com4j.office.XlChartType type,
    @Optional @DefaultValue("-1") float left,
    @Optional @DefaultValue("-1") float top,
    @Optional @DefaultValue("-1") float width,
    @Optional @DefaultValue("-1") float height);


  /**
   * @param numRows Mandatory int parameter.
   * @param numColumns Mandatory int parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @param width Mandatory float parameter.
   * @param height Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(27) //= 0x1b. The runtime will prefer the VTID if present
  @VTID(31)
  com.exceljava.com4j.office.Shape addTable(
    int numRows,
    int numColumns,
    float left,
    float top,
    float width,
    float height);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter left is set to -1.0f</li><li>float parameter top is set to -1.0f</li><li>float parameter width is set to -1.0f</li><li>float parameter height is set to -1.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addSmartArt(layout, -1.0f, -1.0f, -1.0f, -1.0f);
   * </code>
   * </p>
   * @param layout Mandatory com.exceljava.com4j.office.SmartArtLayout parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(28) //= 0x1c. The runtime will prefer the VTID if present
  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {0, 5}, optParamIndex = {1, 2, 3, 4}, javaType = {float.class, float.class, float.class, float.class}, nativeType = {NativeType.Float, NativeType.Float, NativeType.Float, NativeType.Float}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4}, literal = {"-1.0", "-1.0", "-1.0", "-1.0"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.Shape addSmartArt(
    com.exceljava.com4j.office.SmartArtLayout layout);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter top is set to -1.0f</li><li>float parameter width is set to -1.0f</li><li>float parameter height is set to -1.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addSmartArt(layout, left, -1.0f, -1.0f, -1.0f);
   * </code>
   * </p>
   * @param layout Mandatory com.exceljava.com4j.office.SmartArtLayout parameter.
   * @param left Optional parameter. Default value is -1.0f
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(28) //= 0x1c. The runtime will prefer the VTID if present
  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {float.class, float.class, float.class}, nativeType = {NativeType.Float, NativeType.Float, NativeType.Float}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4}, literal = {"-1.0", "-1.0", "-1.0"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.Shape addSmartArt(
    com.exceljava.com4j.office.SmartArtLayout layout,
    @Optional @DefaultValue("-1") float left);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter width is set to -1.0f</li><li>float parameter height is set to -1.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addSmartArt(layout, left, top, -1.0f, -1.0f);
   * </code>
   * </p>
   * @param layout Mandatory com.exceljava.com4j.office.SmartArtLayout parameter.
   * @param left Optional parameter. Default value is -1.0f
   * @param top Optional parameter. Default value is -1.0f
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(28) //= 0x1c. The runtime will prefer the VTID if present
  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {float.class, float.class}, nativeType = {NativeType.Float, NativeType.Float}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_R4}, literal = {"-1.0", "-1.0"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.Shape addSmartArt(
    com.exceljava.com4j.office.SmartArtLayout layout,
    @Optional @DefaultValue("-1") float left,
    @Optional @DefaultValue("-1") float top);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter height is set to -1.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addSmartArt(layout, left, top, width, -1.0f);
   * </code>
   * </p>
   * @param layout Mandatory com.exceljava.com4j.office.SmartArtLayout parameter.
   * @param left Optional parameter. Default value is -1.0f
   * @param top Optional parameter. Default value is -1.0f
   * @param width Optional parameter. Default value is -1.0f
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(28) //= 0x1c. The runtime will prefer the VTID if present
  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {float.class}, nativeType = {NativeType.Float}, variantType = {Variant.Type.VT_R4}, literal = {"-1.0"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.Shape addSmartArt(
    com.exceljava.com4j.office.SmartArtLayout layout,
    @Optional @DefaultValue("-1") float left,
    @Optional @DefaultValue("-1") float top,
    @Optional @DefaultValue("-1") float width);

  /**
   * @param layout Mandatory com.exceljava.com4j.office.SmartArtLayout parameter.
   * @param left Optional parameter. Default value is -1.0f
   * @param top Optional parameter. Default value is -1.0f
   * @param width Optional parameter. Default value is -1.0f
   * @param height Optional parameter. Default value is -1.0f
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(28) //= 0x1c. The runtime will prefer the VTID if present
  @VTID(32)
  com.exceljava.com4j.office.Shape addSmartArt(
    com.exceljava.com4j.office.SmartArtLayout layout,
    @Optional @DefaultValue("-1") float left,
    @Optional @DefaultValue("-1") float top,
    @Optional @DefaultValue("-1") float width,
    @Optional @DefaultValue("-1") float height);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter style is set to -1</li><li>com.exceljava.com4j.office.XlChartType parameter type is set to -1</li><li>float parameter left is set to -1.0f</li><li>float parameter top is set to -1.0f</li><li>float parameter width is set to -1.0f</li><li>float parameter height is set to -1.0f</li><li>boolean parameter newLayout is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart2(-1, -1, -1.0f, -1.0f, -1.0f, -1.0f, false);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(29) //= 0x1d. The runtime will prefer the VTID if present
  @VTID(33)
  @UseDefaultValues(paramIndexMapping = {7}, optParamIndex = {0, 1, 2, 3, 4, 5, 6}, javaType = {int.class, com.exceljava.com4j.office.XlChartType.class, float.class, float.class, float.class, float.class, boolean.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Float, NativeType.Float, NativeType.Float, NativeType.Float, NativeType.VariantBool}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_BOOL}, literal = {"-1", "-1", "-1.0", "-1.0", "-1.0", "-1.0", "false"})
  @ReturnValue(index=7)
  com.exceljava.com4j.office.Shape addChart2();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.XlChartType parameter type is set to -1</li><li>float parameter left is set to -1.0f</li><li>float parameter top is set to -1.0f</li><li>float parameter width is set to -1.0f</li><li>float parameter height is set to -1.0f</li><li>boolean parameter newLayout is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart2(style, -1, -1.0f, -1.0f, -1.0f, -1.0f, false);
   * </code>
   * </p>
   * @param style Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(29) //= 0x1d. The runtime will prefer the VTID if present
  @VTID(33)
  @UseDefaultValues(paramIndexMapping = {0, 7}, optParamIndex = {1, 2, 3, 4, 5, 6}, javaType = {com.exceljava.com4j.office.XlChartType.class, float.class, float.class, float.class, float.class, boolean.class}, nativeType = {NativeType.Int32, NativeType.Float, NativeType.Float, NativeType.Float, NativeType.Float, NativeType.VariantBool}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_BOOL}, literal = {"-1", "-1.0", "-1.0", "-1.0", "-1.0", "false"})
  @ReturnValue(index=7)
  com.exceljava.com4j.office.Shape addChart2(
    @Optional @DefaultValue("-1") int style);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter left is set to -1.0f</li><li>float parameter top is set to -1.0f</li><li>float parameter width is set to -1.0f</li><li>float parameter height is set to -1.0f</li><li>boolean parameter newLayout is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart2(style, type, -1.0f, -1.0f, -1.0f, -1.0f, false);
   * </code>
   * </p>
   * @param style Optional parameter. Default value is -1
   * @param type Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(29) //= 0x1d. The runtime will prefer the VTID if present
  @VTID(33)
  @UseDefaultValues(paramIndexMapping = {0, 1, 7}, optParamIndex = {2, 3, 4, 5, 6}, javaType = {float.class, float.class, float.class, float.class, boolean.class}, nativeType = {NativeType.Float, NativeType.Float, NativeType.Float, NativeType.Float, NativeType.VariantBool}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_BOOL}, literal = {"-1.0", "-1.0", "-1.0", "-1.0", "false"})
  @ReturnValue(index=7)
  com.exceljava.com4j.office.Shape addChart2(
    @Optional @DefaultValue("-1") int style,
    @Optional @DefaultValue("-1") com.exceljava.com4j.office.XlChartType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter top is set to -1.0f</li><li>float parameter width is set to -1.0f</li><li>float parameter height is set to -1.0f</li><li>boolean parameter newLayout is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart2(style, type, left, -1.0f, -1.0f, -1.0f, false);
   * </code>
   * </p>
   * @param style Optional parameter. Default value is -1
   * @param type Optional parameter. Default value is -1
   * @param left Optional parameter. Default value is -1.0f
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(29) //= 0x1d. The runtime will prefer the VTID if present
  @VTID(33)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 7}, optParamIndex = {3, 4, 5, 6}, javaType = {float.class, float.class, float.class, boolean.class}, nativeType = {NativeType.Float, NativeType.Float, NativeType.Float, NativeType.VariantBool}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_BOOL}, literal = {"-1.0", "-1.0", "-1.0", "false"})
  @ReturnValue(index=7)
  com.exceljava.com4j.office.Shape addChart2(
    @Optional @DefaultValue("-1") int style,
    @Optional @DefaultValue("-1") com.exceljava.com4j.office.XlChartType type,
    @Optional @DefaultValue("-1") float left);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter width is set to -1.0f</li><li>float parameter height is set to -1.0f</li><li>boolean parameter newLayout is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart2(style, type, left, top, -1.0f, -1.0f, false);
   * </code>
   * </p>
   * @param style Optional parameter. Default value is -1
   * @param type Optional parameter. Default value is -1
   * @param left Optional parameter. Default value is -1.0f
   * @param top Optional parameter. Default value is -1.0f
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(29) //= 0x1d. The runtime will prefer the VTID if present
  @VTID(33)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 7}, optParamIndex = {4, 5, 6}, javaType = {float.class, float.class, boolean.class}, nativeType = {NativeType.Float, NativeType.Float, NativeType.VariantBool}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_BOOL}, literal = {"-1.0", "-1.0", "false"})
  @ReturnValue(index=7)
  com.exceljava.com4j.office.Shape addChart2(
    @Optional @DefaultValue("-1") int style,
    @Optional @DefaultValue("-1") com.exceljava.com4j.office.XlChartType type,
    @Optional @DefaultValue("-1") float left,
    @Optional @DefaultValue("-1") float top);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter height is set to -1.0f</li><li>boolean parameter newLayout is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart2(style, type, left, top, width, -1.0f, false);
   * </code>
   * </p>
   * @param style Optional parameter. Default value is -1
   * @param type Optional parameter. Default value is -1
   * @param left Optional parameter. Default value is -1.0f
   * @param top Optional parameter. Default value is -1.0f
   * @param width Optional parameter. Default value is -1.0f
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(29) //= 0x1d. The runtime will prefer the VTID if present
  @VTID(33)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 7}, optParamIndex = {5, 6}, javaType = {float.class, boolean.class}, nativeType = {NativeType.Float, NativeType.VariantBool}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_BOOL}, literal = {"-1.0", "false"})
  @ReturnValue(index=7)
  com.exceljava.com4j.office.Shape addChart2(
    @Optional @DefaultValue("-1") int style,
    @Optional @DefaultValue("-1") com.exceljava.com4j.office.XlChartType type,
    @Optional @DefaultValue("-1") float left,
    @Optional @DefaultValue("-1") float top,
    @Optional @DefaultValue("-1") float width);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter newLayout is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart2(style, type, left, top, width, height, false);
   * </code>
   * </p>
   * @param style Optional parameter. Default value is -1
   * @param type Optional parameter. Default value is -1
   * @param left Optional parameter. Default value is -1.0f
   * @param top Optional parameter. Default value is -1.0f
   * @param width Optional parameter. Default value is -1.0f
   * @param height Optional parameter. Default value is -1.0f
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(29) //= 0x1d. The runtime will prefer the VTID if present
  @VTID(33)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 7}, optParamIndex = {6}, javaType = {boolean.class}, nativeType = {NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL}, literal = {"false"})
  @ReturnValue(index=7)
  com.exceljava.com4j.office.Shape addChart2(
    @Optional @DefaultValue("-1") int style,
    @Optional @DefaultValue("-1") com.exceljava.com4j.office.XlChartType type,
    @Optional @DefaultValue("-1") float left,
    @Optional @DefaultValue("-1") float top,
    @Optional @DefaultValue("-1") float width,
    @Optional @DefaultValue("-1") float height);

  /**
   * @param style Optional parameter. Default value is -1
   * @param type Optional parameter. Default value is -1
   * @param left Optional parameter. Default value is -1.0f
   * @param top Optional parameter. Default value is -1.0f
   * @param width Optional parameter. Default value is -1.0f
   * @param height Optional parameter. Default value is -1.0f
   * @param newLayout Optional parameter. Default value is false
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(29) //= 0x1d. The runtime will prefer the VTID if present
  @VTID(33)
  com.exceljava.com4j.office.Shape addChart2(
    @Optional @DefaultValue("-1") int style,
    @Optional @DefaultValue("-1") com.exceljava.com4j.office.XlChartType type,
    @Optional @DefaultValue("-1") float left,
    @Optional @DefaultValue("-1") float top,
    @Optional @DefaultValue("-1") float width,
    @Optional @DefaultValue("-1") float height,
    @Optional @DefaultValue("-1") boolean newLayout);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter width is set to -1.0f</li><li>float parameter height is set to -1.0f</li><li>com.exceljava.com4j.office.MsoPictureCompress parameter compress is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addPicture2(fileName, linkToFile, saveWithDocument, left, top, -1.0f, -1.0f, -1);
   * </code>
   * </p>
   * @param fileName Mandatory java.lang.String parameter.
   * @param linkToFile Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param saveWithDocument Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(30) //= 0x1e. The runtime will prefer the VTID if present
  @VTID(34)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 8}, optParamIndex = {5, 6, 7}, javaType = {float.class, float.class, com.exceljava.com4j.office.MsoPictureCompress.class}, nativeType = {NativeType.Float, NativeType.Float, NativeType.Int32}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_I4}, literal = {"-1.0", "-1.0", "-1"})
  @ReturnValue(index=8)
  com.exceljava.com4j.office.Shape addPicture2(
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
     * <li>float parameter height is set to -1.0f</li><li>com.exceljava.com4j.office.MsoPictureCompress parameter compress is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addPicture2(fileName, linkToFile, saveWithDocument, left, top, width, -1.0f, -1);
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

  @DISPID(30) //= 0x1e. The runtime will prefer the VTID if present
  @VTID(34)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 8}, optParamIndex = {6, 7}, javaType = {float.class, com.exceljava.com4j.office.MsoPictureCompress.class}, nativeType = {NativeType.Float, NativeType.Int32}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_I4}, literal = {"-1.0", "-1"})
  @ReturnValue(index=8)
  com.exceljava.com4j.office.Shape addPicture2(
    java.lang.String fileName,
    com.exceljava.com4j.office.MsoTriState linkToFile,
    com.exceljava.com4j.office.MsoTriState saveWithDocument,
    float left,
    float top,
    @Optional @DefaultValue("-1") float width);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.MsoPictureCompress parameter compress is set to -1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addPicture2(fileName, linkToFile, saveWithDocument, left, top, width, height, -1);
   * </code>
   * </p>
   * @param fileName Mandatory java.lang.String parameter.
   * @param linkToFile Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param saveWithDocument Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @param width Optional parameter. Default value is -1.0f
   * @param height Optional parameter. Default value is -1.0f
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(30) //= 0x1e. The runtime will prefer the VTID if present
  @VTID(34)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 8}, optParamIndex = {7}, javaType = {com.exceljava.com4j.office.MsoPictureCompress.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-1"})
  @ReturnValue(index=8)
  com.exceljava.com4j.office.Shape addPicture2(
    java.lang.String fileName,
    com.exceljava.com4j.office.MsoTriState linkToFile,
    com.exceljava.com4j.office.MsoTriState saveWithDocument,
    float left,
    float top,
    @Optional @DefaultValue("-1") float width,
    @Optional @DefaultValue("-1") float height);

  /**
   * @param fileName Mandatory java.lang.String parameter.
   * @param linkToFile Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param saveWithDocument Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @param width Optional parameter. Default value is -1.0f
   * @param height Optional parameter. Default value is -1.0f
   * @param compress Optional parameter. Default value is -1
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(30) //= 0x1e. The runtime will prefer the VTID if present
  @VTID(34)
  com.exceljava.com4j.office.Shape addPicture2(
    java.lang.String fileName,
    com.exceljava.com4j.office.MsoTriState linkToFile,
    com.exceljava.com4j.office.MsoTriState saveWithDocument,
    float left,
    float top,
    @Optional @DefaultValue("-1") float width,
    @Optional @DefaultValue("-1") float height,
    @Optional @DefaultValue("-1") com.exceljava.com4j.office.MsoPictureCompress compress);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter width is set to -1.0f</li><li>float parameter height is set to -1.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * add3DModel(fileName, linkToFile, saveWithDocument, left, top, -1.0f, -1.0f);
   * </code>
   * </p>
   * @param fileName Mandatory java.lang.String parameter.
   * @param linkToFile Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param saveWithDocument Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(31) //= 0x1f. The runtime will prefer the VTID if present
  @VTID(35)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 7}, optParamIndex = {5, 6}, javaType = {float.class, float.class}, nativeType = {NativeType.Float, NativeType.Float}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_R4}, literal = {"-1.0", "-1.0"})
  @ReturnValue(index=7)
  com.exceljava.com4j.office.Shape add3DModel(
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
   * add3DModel(fileName, linkToFile, saveWithDocument, left, top, width, -1.0f);
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

  @DISPID(31) //= 0x1f. The runtime will prefer the VTID if present
  @VTID(35)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 7}, optParamIndex = {6}, javaType = {float.class}, nativeType = {NativeType.Float}, variantType = {Variant.Type.VT_R4}, literal = {"-1.0"})
  @ReturnValue(index=7)
  com.exceljava.com4j.office.Shape add3DModel(
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

  @DISPID(31) //= 0x1f. The runtime will prefer the VTID if present
  @VTID(35)
  com.exceljava.com4j.office.Shape add3DModel(
    java.lang.String fileName,
    com.exceljava.com4j.office.MsoTriState linkToFile,
    com.exceljava.com4j.office.MsoTriState saveWithDocument,
    float left,
    float top,
    @Optional @DefaultValue("-1") float width,
    @Optional @DefaultValue("-1") float height);


  // Properties:
}
