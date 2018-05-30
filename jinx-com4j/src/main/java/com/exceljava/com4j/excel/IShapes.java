package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002443A-0001-0000-C000-000000000046}")
public interface IShapes extends Com4jObject,Iterable<Com4jObject> {
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
   * @param type Mandatory com.exceljava.com4j.office.MsoCalloutType parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @param width Mandatory float parameter.
   * @param height Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(14)
  com.exceljava.com4j.excel.Shape addCallout(
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
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(15)
  com.exceljava.com4j.excel.Shape addConnector(
    com.exceljava.com4j.office.MsoConnectorType type,
    float beginX,
    float beginY,
    float endX,
    float endY);


  /**
   * @param safeArrayOfPoints Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(16)
  com.exceljava.com4j.excel.Shape addCurve(
    @MarshalAs(NativeType.VARIANT) java.lang.Object safeArrayOfPoints);


  /**
   * @param orientation Mandatory com.exceljava.com4j.office.MsoTextOrientation parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @param width Mandatory float parameter.
   * @param height Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(17)
  com.exceljava.com4j.excel.Shape addLabel(
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
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(18)
  com.exceljava.com4j.excel.Shape addLine(
    float beginX,
    float beginY,
    float endX,
    float endY);


  /**
   * @param filename Mandatory java.lang.String parameter.
   * @param linkToFile Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param saveWithDocument Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @param width Mandatory float parameter.
   * @param height Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(19)
  com.exceljava.com4j.excel.Shape addPicture(
    java.lang.String filename,
    com.exceljava.com4j.office.MsoTriState linkToFile,
    com.exceljava.com4j.office.MsoTriState saveWithDocument,
    float left,
    float top,
    float width,
    float height);


  /**
   * @param safeArrayOfPoints Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(20)
  com.exceljava.com4j.excel.Shape addPolyline(
    @MarshalAs(NativeType.VARIANT) java.lang.Object safeArrayOfPoints);


  /**
   * @param type Mandatory com.exceljava.com4j.office.MsoAutoShapeType parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @param width Mandatory float parameter.
   * @param height Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(21)
  com.exceljava.com4j.excel.Shape addShape(
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
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(22)
  com.exceljava.com4j.excel.Shape addTextEffect(
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
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(23)
  com.exceljava.com4j.excel.Shape addTextbox(
    com.exceljava.com4j.office.MsoTextOrientation orientation,
    float left,
    float top,
    float width,
    float height);


  /**
   * @param editingType Mandatory com.exceljava.com4j.office.MsoEditingType parameter.
   * @param x1 Mandatory float parameter.
   * @param y1 Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.FreeformBuilder
   */

  @VTID(24)
  com.exceljava.com4j.excel.FreeformBuilder buildFreeform(
    com.exceljava.com4j.office.MsoEditingType editingType,
    float x1,
    float y1);


  /**
   * <p>
   * Getter method for the COM property "Range"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ShapeRange
   */

  @VTID(25)
  com.exceljava.com4j.excel.ShapeRange getRange(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   */

  @VTID(26)
  void selectAll();


  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlFormControl parameter.
   * @param left Mandatory int parameter.
   * @param top Mandatory int parameter.
   * @param width Mandatory int parameter.
   * @param height Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(27)
  com.exceljava.com4j.excel.Shape addFormControl(
    com.exceljava.com4j.excel.XlFormControl type,
    int left,
    int top,
    int width,
    int height);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter classType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter filename is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter link is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter displayAsIcon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter left is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addOLEObject(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(28)
  @UseDefaultValues(paramIndexMapping = {11}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Shape addOLEObject();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter filename is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter link is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter displayAsIcon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter left is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addOLEObject(classType, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param classType Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(28)
  @UseDefaultValues(paramIndexMapping = {0, 11}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Shape addOLEObject(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object classType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter link is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter displayAsIcon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter left is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addOLEObject(classType, filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param classType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(28)
  @UseDefaultValues(paramIndexMapping = {0, 1, 11}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Shape addOLEObject(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object classType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter displayAsIcon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter left is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addOLEObject(classType, filename, link, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param classType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(28)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 11}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Shape addOLEObject(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object classType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter iconFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter left is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addOLEObject(classType, filename, link, displayAsIcon, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param classType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(28)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 11}, optParamIndex = {4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Shape addOLEObject(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object classType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayAsIcon);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter iconIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter left is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addOLEObject(classType, filename, link, displayAsIcon, iconFileName, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param classType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(28)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 11}, optParamIndex = {5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Shape addOLEObject(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object classType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayAsIcon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconFileName);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter left is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addOLEObject(classType, filename, link, displayAsIcon, iconFileName, iconIndex, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param classType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(28)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 11}, optParamIndex = {6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Shape addOLEObject(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object classType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayAsIcon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconIndex);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter left is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addOLEObject(classType, filename, link, displayAsIcon, iconFileName, iconIndex, iconLabel, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param classType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconLabel Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(28)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 11}, optParamIndex = {7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Shape addOLEObject(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object classType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayAsIcon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconLabel);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addOLEObject(classType, filename, link, displayAsIcon, iconFileName, iconIndex, iconLabel, left, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param classType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconLabel Optional parameter. Default value is com4j.Variant.getMissing()
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(28)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 11}, optParamIndex = {8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Shape addOLEObject(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object classType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayAsIcon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconLabel,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addOLEObject(classType, filename, link, displayAsIcon, iconFileName, iconIndex, iconLabel, left, top, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param classType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconLabel Optional parameter. Default value is com4j.Variant.getMissing()
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @param top Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(28)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 11}, optParamIndex = {9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Shape addOLEObject(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object classType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayAsIcon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconLabel,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object top);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addOLEObject(classType, filename, link, displayAsIcon, iconFileName, iconIndex, iconLabel, left, top, width, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param classType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconLabel Optional parameter. Default value is com4j.Variant.getMissing()
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @param top Optional parameter. Default value is com4j.Variant.getMissing()
   * @param width Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(28)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 11}, optParamIndex = {10}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=11)
  com.exceljava.com4j.excel.Shape addOLEObject(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object classType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayAsIcon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconLabel,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object top,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object width);

  /**
   * @param classType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconLabel Optional parameter. Default value is com4j.Variant.getMissing()
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @param top Optional parameter. Default value is com4j.Variant.getMissing()
   * @param width Optional parameter. Default value is com4j.Variant.getMissing()
   * @param height Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(28)
  com.exceljava.com4j.excel.Shape addOLEObject(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object classType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayAsIcon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconLabel,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object top,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object width,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object height);


  /**
   * @param type Mandatory com.exceljava.com4j.office.MsoDiagramType parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @param width Mandatory float parameter.
   * @param height Mandatory float parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(29)
  com.exceljava.com4j.excel.Shape addDiagram(
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
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(30)
  com.exceljava.com4j.excel.Shape addCanvas(
    float left,
    float top,
    float width,
    float height);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter xlChartType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter left is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(31)
  @UseDefaultValues(paramIndexMapping = {5}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.Shape addChart();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter left is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart(xlChartType, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param xlChartType Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(31)
  @UseDefaultValues(paramIndexMapping = {0, 5}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.Shape addChart(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object xlChartType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart(xlChartType, left, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param xlChartType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(31)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.Shape addChart(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object xlChartType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart(xlChartType, left, top, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param xlChartType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @param top Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(31)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.Shape addChart(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object xlChartType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object top);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart(xlChartType, left, top, width, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param xlChartType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @param top Optional parameter. Default value is com4j.Variant.getMissing()
   * @param width Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(31)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.Shape addChart(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object xlChartType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object top,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object width);

  /**
   * @param xlChartType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @param top Optional parameter. Default value is com4j.Variant.getMissing()
   * @param width Optional parameter. Default value is com4j.Variant.getMissing()
   * @param height Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(31)
  com.exceljava.com4j.excel.Shape addChart(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object xlChartType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object top,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object width,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object height);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter left is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addSmartArt(layout, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param layout Mandatory com.exceljava.com4j.office.SmartArtLayout parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {0, 5}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.Shape addSmartArt(
    com.exceljava.com4j.office.SmartArtLayout layout);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addSmartArt(layout, left, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param layout Mandatory com.exceljava.com4j.office.SmartArtLayout parameter.
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.Shape addSmartArt(
    com.exceljava.com4j.office.SmartArtLayout layout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addSmartArt(layout, left, top, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param layout Mandatory com.exceljava.com4j.office.SmartArtLayout parameter.
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @param top Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.Shape addSmartArt(
    com.exceljava.com4j.office.SmartArtLayout layout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object top);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addSmartArt(layout, left, top, width, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param layout Mandatory com.exceljava.com4j.office.SmartArtLayout parameter.
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @param top Optional parameter. Default value is com4j.Variant.getMissing()
   * @param width Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.excel.Shape addSmartArt(
    com.exceljava.com4j.office.SmartArtLayout layout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object top,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object width);

  /**
   * @param layout Mandatory com.exceljava.com4j.office.SmartArtLayout parameter.
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @param top Optional parameter. Default value is com4j.Variant.getMissing()
   * @param width Optional parameter. Default value is com4j.Variant.getMissing()
   * @param height Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(32)
  com.exceljava.com4j.excel.Shape addSmartArt(
    com.exceljava.com4j.office.SmartArtLayout layout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object top,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object width,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object height);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter style is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter xlChartType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter left is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter newLayout is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart2(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(33)
  @UseDefaultValues(paramIndexMapping = {7}, optParamIndex = {0, 1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.Shape addChart2();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter xlChartType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter left is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter newLayout is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart2(style, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param style Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(33)
  @UseDefaultValues(paramIndexMapping = {0, 7}, optParamIndex = {1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.Shape addChart2(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object style);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter left is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter newLayout is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart2(style, xlChartType, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param style Optional parameter. Default value is com4j.Variant.getMissing()
   * @param xlChartType Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(33)
  @UseDefaultValues(paramIndexMapping = {0, 1, 7}, optParamIndex = {2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.Shape addChart2(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object style,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object xlChartType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter top is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter newLayout is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart2(style, xlChartType, left, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param style Optional parameter. Default value is com4j.Variant.getMissing()
   * @param xlChartType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(33)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 7}, optParamIndex = {3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.Shape addChart2(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object style,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object xlChartType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter width is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter newLayout is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart2(style, xlChartType, left, top, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param style Optional parameter. Default value is com4j.Variant.getMissing()
   * @param xlChartType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @param top Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(33)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 7}, optParamIndex = {4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.Shape addChart2(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object style,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object xlChartType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object top);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter height is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter newLayout is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart2(style, xlChartType, left, top, width, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param style Optional parameter. Default value is com4j.Variant.getMissing()
   * @param xlChartType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @param top Optional parameter. Default value is com4j.Variant.getMissing()
   * @param width Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(33)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 7}, optParamIndex = {5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.Shape addChart2(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object style,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object xlChartType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object top,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object width);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter newLayout is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addChart2(style, xlChartType, left, top, width, height, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param style Optional parameter. Default value is com4j.Variant.getMissing()
   * @param xlChartType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @param top Optional parameter. Default value is com4j.Variant.getMissing()
   * @param width Optional parameter. Default value is com4j.Variant.getMissing()
   * @param height Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(33)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 7}, optParamIndex = {6}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=7)
  com.exceljava.com4j.excel.Shape addChart2(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object style,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object xlChartType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object top,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object width,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object height);

  /**
   * @param style Optional parameter. Default value is com4j.Variant.getMissing()
   * @param xlChartType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param left Optional parameter. Default value is com4j.Variant.getMissing()
   * @param top Optional parameter. Default value is com4j.Variant.getMissing()
   * @param width Optional parameter. Default value is com4j.Variant.getMissing()
   * @param height Optional parameter. Default value is com4j.Variant.getMissing()
   * @param newLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(33)
  com.exceljava.com4j.excel.Shape addChart2(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object style,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object xlChartType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object left,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object top,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object width,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object height,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object newLayout);


  /**
   * @param filename Mandatory java.lang.String parameter.
   * @param linkToFile Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param saveWithDocument Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   * @param left Mandatory float parameter.
   * @param top Mandatory float parameter.
   * @param width Mandatory float parameter.
   * @param height Mandatory float parameter.
   * @param compress Mandatory com.exceljava.com4j.office.MsoPictureCompress parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(34)
  com.exceljava.com4j.excel.Shape addPicture2(
    java.lang.String filename,
    com.exceljava.com4j.office.MsoTriState linkToFile,
    com.exceljava.com4j.office.MsoTriState saveWithDocument,
    float left,
    float top,
    float width,
    float height,
    com.exceljava.com4j.office.MsoPictureCompress compress);


  // Properties:
}
