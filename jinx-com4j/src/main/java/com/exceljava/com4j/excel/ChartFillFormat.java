package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ChartFillFormat extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   */

  @DISPID(148)
  @PropGet
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   */

  @DISPID(149)
  @PropGet
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   */

  @DISPID(150)
  @PropGet
  com4j.Com4jObject getParent();


  /**
   * @param style Mandatory com.exceljava.com4j.office.MsoGradientStyle parameter.
   * @param variant Mandatory int parameter.
   * @param degree Mandatory float parameter.
   */

  @DISPID(1621)
  void oneColorGradient(
    com.exceljava.com4j.office.MsoGradientStyle style,
    int variant,
    float degree);


  /**
   * @param style Mandatory com.exceljava.com4j.office.MsoGradientStyle parameter.
   * @param variant Mandatory int parameter.
   */

  @DISPID(1624)
  void twoColorGradient(
    com.exceljava.com4j.office.MsoGradientStyle style,
    int variant);


  /**
   * @param presetTexture Mandatory com.exceljava.com4j.office.MsoPresetTexture parameter.
   */

  @DISPID(1625)
  void presetTextured(
    com.exceljava.com4j.office.MsoPresetTexture presetTexture);


  /**
   */

  @DISPID(1627)
  void solid();


  /**
   * @param pattern Mandatory com.exceljava.com4j.office.MsoPatternType parameter.
   */

  @DISPID(1628)
  void patterned(
    com.exceljava.com4j.office.MsoPatternType pattern);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter pictureFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pictureFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pictureStackUnit is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter picturePlacement is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * userPicture(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1629)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void userPicture();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter pictureFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pictureStackUnit is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter picturePlacement is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * userPicture(pictureFile, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param pictureFile Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1629)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void userPicture(
    @Optional java.lang.Object pictureFile);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter pictureStackUnit is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter picturePlacement is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * userPicture(pictureFile, pictureFormat, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param pictureFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pictureFormat Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1629)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void userPicture(
    @Optional java.lang.Object pictureFile,
    @Optional java.lang.Object pictureFormat);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter picturePlacement is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * userPicture(pictureFile, pictureFormat, pictureStackUnit, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param pictureFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pictureFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pictureStackUnit Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1629)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void userPicture(
    @Optional java.lang.Object pictureFile,
    @Optional java.lang.Object pictureFormat,
    @Optional java.lang.Object pictureStackUnit);

  /**
   * @param pictureFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pictureFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pictureStackUnit Optional parameter. Default value is com4j.Variant.getMissing()
   * @param picturePlacement Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1629)
  void userPicture(
    @Optional java.lang.Object pictureFile,
    @Optional java.lang.Object pictureFormat,
    @Optional java.lang.Object pictureStackUnit,
    @Optional java.lang.Object picturePlacement);


  /**
   * @param textureFile Mandatory java.lang.String parameter.
   */

  @DISPID(1634)
  void userTextured(
    java.lang.String textureFile);


  /**
   * @param style Mandatory com.exceljava.com4j.office.MsoGradientStyle parameter.
   * @param variant Mandatory int parameter.
   * @param presetGradientType Mandatory com.exceljava.com4j.office.MsoPresetGradientType parameter.
   */

  @DISPID(1636)
  void presetGradient(
    com.exceljava.com4j.office.MsoGradientStyle style,
    int variant,
    com.exceljava.com4j.office.MsoPresetGradientType presetGradientType);


  /**
   * <p>
   * Getter method for the COM property "BackColor"
   * </p>
   */

  @DISPID(1638)
  @PropGet
  com.exceljava.com4j.excel.ChartColorFormat getBackColor();


  /**
   * <p>
   * Getter method for the COM property "ForeColor"
   * </p>
   */

  @DISPID(1639)
  @PropGet
  com.exceljava.com4j.excel.ChartColorFormat getForeColor();


  /**
   * <p>
   * Getter method for the COM property "GradientColorType"
   * </p>
   */

  @DISPID(1640)
  @PropGet
  com.exceljava.com4j.office.MsoGradientColorType getGradientColorType();


  /**
   * <p>
   * Getter method for the COM property "GradientDegree"
   * </p>
   */

  @DISPID(1641)
  @PropGet
  float getGradientDegree();


  /**
   * <p>
   * Getter method for the COM property "GradientStyle"
   * </p>
   */

  @DISPID(1642)
  @PropGet
  com.exceljava.com4j.office.MsoGradientStyle getGradientStyle();


  /**
   * <p>
   * Getter method for the COM property "GradientVariant"
   * </p>
   */

  @DISPID(1643)
  @PropGet
  int getGradientVariant();


  /**
   * <p>
   * Getter method for the COM property "Pattern"
   * </p>
   */

  @DISPID(95)
  @PropGet
  com.exceljava.com4j.office.MsoPatternType getPattern();


  /**
   * <p>
   * Getter method for the COM property "PresetGradientType"
   * </p>
   */

  @DISPID(1637)
  @PropGet
  com.exceljava.com4j.office.MsoPresetGradientType getPresetGradientType();


  /**
   * <p>
   * Getter method for the COM property "PresetTexture"
   * </p>
   */

  @DISPID(1626)
  @PropGet
  com.exceljava.com4j.office.MsoPresetTexture getPresetTexture();


  /**
   * <p>
   * Getter method for the COM property "TextureName"
   * </p>
   */

  @DISPID(1644)
  @PropGet
  java.lang.String getTextureName();


  /**
   * <p>
   * Getter method for the COM property "TextureType"
   * </p>
   */

  @DISPID(1645)
  @PropGet
  com.exceljava.com4j.office.MsoTextureType getTextureType();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   */

  @DISPID(108)
  @PropGet
  com.exceljava.com4j.office.MsoFillType getType();


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   */

  @DISPID(558)
  @PropGet
  com.exceljava.com4j.office.MsoTriState getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(558)
  @PropPut
  void setVisible(
    com.exceljava.com4j.office.MsoTriState rhs);


  // Properties:
}
