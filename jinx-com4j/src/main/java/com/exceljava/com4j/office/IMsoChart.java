package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C1709-0000-0000-C000-000000000046}")
public interface IMsoChart extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(150) //= 0x96. The runtime will prefer the VTID if present
  @VTID(7)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Setter method for the COM property "HasTitle"
   * </p>
   * @param fTitle Mandatory boolean parameter.
   */

  @DISPID(1610743809) //= 0x60020001. The runtime will prefer the VTID if present
  @VTID(8)
  void setHasTitle(
    boolean fTitle);


  /**
   * <p>
   * Getter method for the COM property "HasTitle"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743809) //= 0x60020001. The runtime will prefer the VTID if present
  @VTID(9)
  boolean getHasTitle();


  /**
   * <p>
   * Getter method for the COM property "ChartTitle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChartTitle
   */

  @DISPID(1610743811) //= 0x60020003. The runtime will prefer the VTID if present
  @VTID(10)
  com.exceljava.com4j.office.IMsoChartTitle getChartTitle();


  /**
   * <p>
   * Getter method for the COM property "DepthPercent"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743812) //= 0x60020004. The runtime will prefer the VTID if present
  @VTID(11)
  int getDepthPercent();


  /**
   * <p>
   * Setter method for the COM property "DepthPercent"
   * </p>
   * @param pwDepthPercent Mandatory int parameter.
   */

  @DISPID(1610743812) //= 0x60020004. The runtime will prefer the VTID if present
  @VTID(12)
  void setDepthPercent(
    int pwDepthPercent);


  /**
   * <p>
   * Getter method for the COM property "Elevation"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743814) //= 0x60020006. The runtime will prefer the VTID if present
  @VTID(13)
  int getElevation();


  /**
   * <p>
   * Setter method for the COM property "Elevation"
   * </p>
   * @param pwElevation Mandatory int parameter.
   */

  @DISPID(1610743814) //= 0x60020006. The runtime will prefer the VTID if present
  @VTID(14)
  void setElevation(
    int pwElevation);


  /**
   * <p>
   * Getter method for the COM property "GapDepth"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743816) //= 0x60020008. The runtime will prefer the VTID if present
  @VTID(15)
  int getGapDepth();


  /**
   * <p>
   * Setter method for the COM property "GapDepth"
   * </p>
   * @param pwGapDepth Mandatory int parameter.
   */

  @DISPID(1610743816) //= 0x60020008. The runtime will prefer the VTID if present
  @VTID(16)
  void setGapDepth(
    int pwGapDepth);


  /**
   * <p>
   * Getter method for the COM property "HeightPercent"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743818) //= 0x6002000a. The runtime will prefer the VTID if present
  @VTID(17)
  int getHeightPercent();


  /**
   * <p>
   * Setter method for the COM property "HeightPercent"
   * </p>
   * @param pwHeightPercent Mandatory int parameter.
   */

  @DISPID(1610743818) //= 0x6002000a. The runtime will prefer the VTID if present
  @VTID(18)
  void setHeightPercent(
    int pwHeightPercent);


  /**
   * <p>
   * Getter method for the COM property "Perspective"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610743820) //= 0x6002000c. The runtime will prefer the VTID if present
  @VTID(19)
  int getPerspective();


  /**
   * <p>
   * Setter method for the COM property "Perspective"
   * </p>
   * @param pwPerspective Mandatory int parameter.
   */

  @DISPID(1610743820) //= 0x6002000c. The runtime will prefer the VTID if present
  @VTID(20)
  void setPerspective(
    int pwPerspective);


  /**
   * <p>
   * Getter method for the COM property "RightAngleAxes"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1610743822) //= 0x6002000e. The runtime will prefer the VTID if present
  @VTID(21)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getRightAngleAxes();


  /**
   * <p>
   * Setter method for the COM property "RightAngleAxes"
   * </p>
   * @param pvarRightAngleAxes Mandatory java.lang.Object parameter.
   */

  @DISPID(1610743822) //= 0x6002000e. The runtime will prefer the VTID if present
  @VTID(22)
  void setRightAngleAxes(
    @MarshalAs(NativeType.VARIANT) java.lang.Object pvarRightAngleAxes);


  /**
   * <p>
   * Getter method for the COM property "Rotation"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1610743824) //= 0x60020010. The runtime will prefer the VTID if present
  @VTID(23)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getRotation();


  /**
   * <p>
   * Setter method for the COM property "Rotation"
   * </p>
   * @param pvarRotation Mandatory java.lang.Object parameter.
   */

  @DISPID(1610743824) //= 0x60020010. The runtime will prefer the VTID if present
  @VTID(24)
  void setRotation(
    @MarshalAs(NativeType.VARIANT) java.lang.Object pvarRotation);


  /**
   * <p>
   * Setter method for the COM property "DisplayBlanksAs"
   * </p>
   * @param pres Mandatory com.exceljava.com4j.office.XlDisplayBlanksAs parameter.
   */

  @DISPID(1610743826) //= 0x60020012. The runtime will prefer the VTID if present
  @VTID(25)
  void setDisplayBlanksAs(
    com.exceljava.com4j.office.XlDisplayBlanksAs pres);


  /**
   * <p>
   * Getter method for the COM property "DisplayBlanksAs"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.XlDisplayBlanksAs
   */

  @DISPID(1610743826) //= 0x60020012. The runtime will prefer the VTID if present
  @VTID(26)
  com.exceljava.com4j.office.XlDisplayBlanksAs getDisplayBlanksAs();


  /**
   * <p>
   * Setter method for the COM property "ProtectData"
   * </p>
   * @param pres Mandatory boolean parameter.
   */

  @DISPID(1610743828) //= 0x60020014. The runtime will prefer the VTID if present
  @VTID(27)
  void setProtectData(
    boolean pres);


  /**
   * <p>
   * Getter method for the COM property "ProtectData"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743828) //= 0x60020014. The runtime will prefer the VTID if present
  @VTID(28)
  boolean getProtectData();


  /**
   * <p>
   * Setter method for the COM property "ProtectFormatting"
   * </p>
   * @param pres Mandatory boolean parameter.
   */

  @DISPID(1610743830) //= 0x60020016. The runtime will prefer the VTID if present
  @VTID(29)
  void setProtectFormatting(
    boolean pres);


  /**
   * <p>
   * Getter method for the COM property "ProtectFormatting"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743830) //= 0x60020016. The runtime will prefer the VTID if present
  @VTID(30)
  boolean getProtectFormatting();


  /**
   * <p>
   * Setter method for the COM property "ProtectGoalSeek"
   * </p>
   * @param pres Mandatory boolean parameter.
   */

  @DISPID(1610743832) //= 0x60020018. The runtime will prefer the VTID if present
  @VTID(31)
  void setProtectGoalSeek(
    boolean pres);


  /**
   * <p>
   * Getter method for the COM property "ProtectGoalSeek"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743832) //= 0x60020018. The runtime will prefer the VTID if present
  @VTID(32)
  boolean getProtectGoalSeek();


  /**
   * <p>
   * Setter method for the COM property "ProtectSelection"
   * </p>
   * @param pres Mandatory boolean parameter.
   */

  @DISPID(1610743834) //= 0x6002001a. The runtime will prefer the VTID if present
  @VTID(33)
  void setProtectSelection(
    boolean pres);


  /**
   * <p>
   * Getter method for the COM property "ProtectSelection"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743834) //= 0x6002001a. The runtime will prefer the VTID if present
  @VTID(34)
  boolean getProtectSelection();


  /**
   * <p>
   * Setter method for the COM property "ProtectChartObjects"
   * </p>
   * @param pres Mandatory boolean parameter.
   */

  @DISPID(1610743836) //= 0x6002001c. The runtime will prefer the VTID if present
  @VTID(35)
  void setProtectChartObjects(
    boolean pres);


  /**
   * <p>
   * Getter method for the COM property "ProtectChartObjects"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743836) //= 0x6002001c. The runtime will prefer the VTID if present
  @VTID(36)
  boolean getProtectChartObjects();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * unProtect(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1610743838) //= 0x6002001e. The runtime will prefer the VTID if present
  @VTID(37)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void unProtect();

  /**
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743838) //= 0x6002001e. The runtime will prefer the VTID if present
  @VTID(37)
  void unProtect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter drawingObjects is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter contents is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter scenarios is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter userInterfaceOnly is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1610743839) //= 0x6002001f. The runtime will prefer the VTID if present
  @VTID(38)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  void protect();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter drawingObjects is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter contents is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter scenarios is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter userInterfaceOnly is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743839) //= 0x6002001f. The runtime will prefer the VTID if present
  @VTID(38)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter contents is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter scenarios is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter userInterfaceOnly is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, drawingObjects, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743839) //= 0x6002001f. The runtime will prefer the VTID if present
  @VTID(38)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter scenarios is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter userInterfaceOnly is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, drawingObjects, contents, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743839) //= 0x6002001f. The runtime will prefer the VTID if present
  @VTID(38)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter userInterfaceOnly is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, drawingObjects, contents, scenarios, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scenarios Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743839) //= 0x6002001f. The runtime will prefer the VTID if present
  @VTID(38)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scenarios);

  /**
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scenarios Optional parameter. Default value is com4j.Variant.getMissing()
   * @param userInterfaceOnly Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743839) //= 0x6002001f. The runtime will prefer the VTID if present
  @VTID(38)
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scenarios,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object userInterfaceOnly);


  /**
   * <p>
   * Getter method for the COM property "ChartGroups"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter pvarIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varIgallery is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getChartGroups(com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(39)
  @UseDefaultValues(paramIndexMapping = {3}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT_ByRef, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.NO_TYPE, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=3)
  com4j.Com4jObject getChartGroups();

  /**
   * <p>
   * Getter method for the COM property "ChartGroups"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varIgallery is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getChartGroups(pvarIndex, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param pvarIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(39)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=3)
  com4j.Com4jObject getChartGroups(
    @Optional java.lang.Object pvarIndex);

  /**
   * <p>
   * Getter method for the COM property "ChartGroups"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getChartGroups(pvarIndex, varIgallery, 1033);
   * </code>
   * </p>
   * @param pvarIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varIgallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(39)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=3)
  com4j.Com4jObject getChartGroups(
    @Optional java.lang.Object pvarIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varIgallery);

  /**
   * <p>
   * Getter method for the COM property "ChartGroups"
   * </p>
   * @param pvarIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varIgallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(39)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getChartGroups(
    @Optional java.lang.Object pvarIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varIgallery,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * seriesCollection(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(68) //= 0x44. The runtime will prefer the VTID if present
  @VTID(40)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.Dispatch,index=1)
  com4j.Com4jObject seriesCollection();

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(68) //= 0x44. The runtime will prefer the VTID if present
  @VTID(40)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject seriesCollection(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.XlDataLabelsType parameter type is set to 2</li><li>java.lang.Object parameter iMsoLegendKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ApplyDataLabels(2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(151) //= 0x97. The runtime will prefer the VTID if present
  @VTID(41)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {com.exceljava.com4j.office.XlDataLabelsType.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"2", "80020004", "80020004", "80020004"})
  void _ApplyDataLabels();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter iMsoLegendKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ApplyDataLabels(type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   */

  @DISPID(151) //= 0x97. The runtime will prefer the VTID if present
  @VTID(41)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void _ApplyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlDataLabelsType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter autoText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ApplyDataLabels(type, iMsoLegendKey, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param iMsoLegendKey Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(151) //= 0x97. The runtime will prefer the VTID if present
  @VTID(41)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void _ApplyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iMsoLegendKey);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ApplyDataLabels(type, iMsoLegendKey, autoText, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param iMsoLegendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(151) //= 0x97. The runtime will prefer the VTID if present
  @VTID(41)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void _ApplyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iMsoLegendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText);

  /**
   * @param type Optional parameter. Default value is 2
   * @param iMsoLegendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(151) //= 0x97. The runtime will prefer the VTID if present
  @VTID(41)
  void _ApplyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iMsoLegendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines);


  /**
   * <p>
   * Getter method for the COM property "SubType"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(42)
  int getSubType();


  /**
   * <p>
   * Setter method for the COM property "SubType"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(43)
  void setSubType(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(44)
  int getType();


  /**
   * <p>
   * Setter method for the COM property "Type"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(45)
  void setType(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Corners"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoCorners
   */

  @DISPID(79) //= 0x4f. The runtime will prefer the VTID if present
  @VTID(46)
  com.exceljava.com4j.office.IMsoCorners getCorners();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.XlDataLabelsType parameter type is set to 2</li><li>java.lang.Object parameter iMsoLegendKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showSeriesName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showCategoryName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, javaType = {com.exceljava.com4j.office.XlDataLabelsType.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"2", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void applyDataLabels();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter iMsoLegendKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showSeriesName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showCategoryName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlDataLabelsType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter autoText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showSeriesName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showCategoryName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, iMsoLegendKey, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param iMsoLegendKey Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iMsoLegendKey);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showSeriesName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showCategoryName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, iMsoLegendKey, autoText, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param iMsoLegendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iMsoLegendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter showSeriesName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showCategoryName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, iMsoLegendKey, autoText, hasLeaderLines, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param iMsoLegendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iMsoLegendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter showCategoryName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param iMsoLegendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showSeriesName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iMsoLegendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showSeriesName);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param iMsoLegendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showSeriesName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showCategoryName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iMsoLegendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showSeriesName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showCategoryName);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param iMsoLegendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showSeriesName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showCategoryName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showValue Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iMsoLegendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showSeriesName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showCategoryName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showValue);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param iMsoLegendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showSeriesName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showCategoryName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showValue Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showPercentage Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iMsoLegendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showSeriesName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showCategoryName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showValue,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showPercentage);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param iMsoLegendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showSeriesName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showCategoryName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showValue Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showPercentage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showBubbleSize Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iMsoLegendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showSeriesName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showCategoryName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showValue,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showPercentage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showBubbleSize);

  /**
   * @param type Optional parameter. Default value is 2
   * @param iMsoLegendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showSeriesName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showCategoryName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showValue Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showPercentage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showBubbleSize Optional parameter. Default value is com4j.Variant.getMissing()
   * @param separator Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(47)
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.office.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iMsoLegendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showSeriesName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showCategoryName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showValue,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showPercentage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showBubbleSize,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object separator);


  /**
   * <p>
   * Getter method for the COM property "ChartType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.XlChartType
   */

  @DISPID(1400) //= 0x578. The runtime will prefer the VTID if present
  @VTID(48)
  com.exceljava.com4j.office.XlChartType getChartType();


  /**
   * <p>
   * Setter method for the COM property "ChartType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.XlChartType parameter.
   */

  @DISPID(1400) //= 0x578. The runtime will prefer the VTID if present
  @VTID(49)
  void setChartType(
    com.exceljava.com4j.office.XlChartType rhs);


  /**
   * <p>
   * Getter method for the COM property "HasDataTable"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1396) //= 0x574. The runtime will prefer the VTID if present
  @VTID(50)
  boolean getHasDataTable();


  /**
   * <p>
   * Setter method for the COM property "HasDataTable"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1396) //= 0x574. The runtime will prefer the VTID if present
  @VTID(51)
  void setHasDataTable(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter typeName is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyCustomType(chartType, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param chartType Mandatory com.exceljava.com4j.office.XlChartType parameter.
   */

  @DISPID(1401) //= 0x579. The runtime will prefer the VTID if present
  @VTID(52)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void applyCustomType(
    com.exceljava.com4j.office.XlChartType chartType);

  /**
   * @param chartType Mandatory com.exceljava.com4j.office.XlChartType parameter.
   * @param typeName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1401) //= 0x579. The runtime will prefer the VTID if present
  @VTID(52)
  void applyCustomType(
    com.exceljava.com4j.office.XlChartType chartType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object typeName);


  /**
   * @param x Mandatory int parameter.
   * @param y Mandatory int parameter.
   * @param elementID Mandatory Holder<Integer> parameter.
   * @param arg1 Mandatory Holder<Integer> parameter.
   * @param arg2 Mandatory Holder<Integer> parameter.
   */

  @DISPID(1409) //= 0x581. The runtime will prefer the VTID if present
  @VTID(53)
  void getChartElement(
    int x,
    int y,
    Holder<Integer> elementID,
    Holder<Integer> arg1,
    Holder<Integer> arg2);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter plotBy is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setSourceData(source, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param source Mandatory java.lang.String parameter.
   */

  @DISPID(1413) //= 0x585. The runtime will prefer the VTID if present
  @VTID(54)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setSourceData(
    java.lang.String source);

  /**
   * @param source Mandatory java.lang.String parameter.
   * @param plotBy Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1413) //= 0x585. The runtime will prefer the VTID if present
  @VTID(54)
  void setSourceData(
    java.lang.String source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object plotBy);


  /**
   * <p>
   * Getter method for the COM property "PlotBy"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.XlRowCol
   */

  @DISPID(202) //= 0xca. The runtime will prefer the VTID if present
  @VTID(55)
  com.exceljava.com4j.office.XlRowCol getPlotBy();


  /**
   * <p>
   * Setter method for the COM property "PlotBy"
   * </p>
   * @param plotBy Mandatory com.exceljava.com4j.office.XlRowCol parameter.
   */

  @DISPID(202) //= 0xca. The runtime will prefer the VTID if present
  @VTID(56)
  void setPlotBy(
    com.exceljava.com4j.office.XlRowCol plotBy);


  /**
   * <p>
   * Getter method for the COM property "HasLegend"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(53) //= 0x35. The runtime will prefer the VTID if present
  @VTID(57)
  boolean getHasLegend();


  /**
   * <p>
   * Setter method for the COM property "HasLegend"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(53) //= 0x35. The runtime will prefer the VTID if present
  @VTID(58)
  void setHasLegend(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Legend"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoLegend
   */

  @DISPID(84) //= 0x54. The runtime will prefer the VTID if present
  @VTID(59)
  com.exceljava.com4j.office.IMsoLegend getLegend();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.office.XlAxisGroup parameter axisGroup is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * axes(com4j.Variant.getMissing(), 1);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610743861) //= 0x60020035. The runtime will prefer the VTID if present
  @VTID(60)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, com.exceljava.com4j.office.XlAxisGroup.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject axes();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.XlAxisGroup parameter axisGroup is set to 1</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * axes(type, 1);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610743861) //= 0x60020035. The runtime will prefer the VTID if present
  @VTID(60)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {com.exceljava.com4j.office.XlAxisGroup.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject axes(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type);

  /**
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param axisGroup Optional parameter. Default value is 1
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610743861) //= 0x60020035. The runtime will prefer the VTID if present
  @VTID(60)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject axes(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.office.XlAxisGroup axisGroup);


  /**
   * <p>
   * Setter method for the COM property "HasAxis"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter axisType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter axisGroup is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setHasAxis(com4j.Variant.getMissing(), com4j.Variant.getMissing(), pval);
   * </code>
   * </p>
   * @param pval Mandatory java.lang.Object parameter.
   */

  @DISPID(1610743862) //= 0x60020036. The runtime will prefer the VTID if present
  @VTID(61)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void setHasAxis(
    @MarshalAs(NativeType.VARIANT) java.lang.Object pval);

  /**
   * <p>
   * Setter method for the COM property "HasAxis"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter axisGroup is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setHasAxis(axisType, com4j.Variant.getMissing(), pval);
   * </code>
   * </p>
   * @param axisType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pval Mandatory java.lang.Object parameter.
   */

  @DISPID(1610743862) //= 0x60020036. The runtime will prefer the VTID if present
  @VTID(61)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setHasAxis(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object axisType,
    @MarshalAs(NativeType.VARIANT) java.lang.Object pval);

  /**
   * <p>
   * Setter method for the COM property "HasAxis"
   * </p>
   * @param axisType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param axisGroup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pval Mandatory java.lang.Object parameter.
   */

  @DISPID(1610743862) //= 0x60020036. The runtime will prefer the VTID if present
  @VTID(61)
  void setHasAxis(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object axisType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object axisGroup,
    @MarshalAs(NativeType.VARIANT) java.lang.Object pval);


  /**
   * <p>
   * Getter method for the COM property "HasAxis"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter axisType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter axisGroup is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getHasAxis(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1610743862) //= 0x60020036. The runtime will prefer the VTID if present
  @VTID(62)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object getHasAxis();

  /**
   * <p>
   * Getter method for the COM property "HasAxis"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter axisGroup is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getHasAxis(axisType, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param axisType Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1610743862) //= 0x60020036. The runtime will prefer the VTID if present
  @VTID(62)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object getHasAxis(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object axisType);

  /**
   * <p>
   * Getter method for the COM property "HasAxis"
   * </p>
   * @param axisType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param axisGroup Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1610743862) //= 0x60020036. The runtime will prefer the VTID if present
  @VTID(62)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getHasAxis(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object axisType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object axisGroup);


  /**
   * <p>
   * Getter method for the COM property "Walls"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter fBackWall is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getWalls(false);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoWalls
   */

  @DISPID(1610743864) //= 0x60020038. The runtime will prefer the VTID if present
  @VTID(63)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {boolean.class}, nativeType = {NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL}, literal = {"false"})
  @ReturnValue(index=1)
  com.exceljava.com4j.office.IMsoWalls getWalls();

  /**
   * <p>
   * Getter method for the COM property "Walls"
   * </p>
   * @param fBackWall Optional parameter. Default value is false
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoWalls
   */

  @DISPID(1610743864) //= 0x60020038. The runtime will prefer the VTID if present
  @VTID(63)
  com.exceljava.com4j.office.IMsoWalls getWalls(
    @Optional @DefaultValue("-1") boolean fBackWall);


  /**
   * <p>
   * Getter method for the COM property "Floor"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoFloor
   */

  @DISPID(1610743865) //= 0x60020039. The runtime will prefer the VTID if present
  @VTID(64)
  com.exceljava.com4j.office.IMsoFloor getFloor();


  /**
   * <p>
   * Getter method for the COM property "PlotArea"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoPlotArea
   */

  @DISPID(1610743866) //= 0x6002003a. The runtime will prefer the VTID if present
  @VTID(65)
  com.exceljava.com4j.office.IMsoPlotArea getPlotArea();


  /**
   * <p>
   * Getter method for the COM property "PlotVisibleOnly"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(92) //= 0x5c. The runtime will prefer the VTID if present
  @VTID(66)
  boolean getPlotVisibleOnly();


  /**
   * <p>
   * Setter method for the COM property "PlotVisibleOnly"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(92) //= 0x5c. The runtime will prefer the VTID if present
  @VTID(67)
  void setPlotVisibleOnly(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ChartArea"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChartArea
   */

  @DISPID(1610743869) //= 0x6002003d. The runtime will prefer the VTID if present
  @VTID(68)
  com.exceljava.com4j.office.IMsoChartArea getChartArea();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varFormat is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * autoFormat(rGallery, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param rGallery Mandatory int parameter.
   */

  @DISPID(1610743870) //= 0x6002003e. The runtime will prefer the VTID if present
  @VTID(69)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void autoFormat(
    int rGallery);

  /**
   * @param rGallery Mandatory int parameter.
   * @param varFormat Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743870) //= 0x6002003e. The runtime will prefer the VTID if present
  @VTID(69)
  void autoFormat(
    int rGallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varFormat);


  /**
   * <p>
   * Getter method for the COM property "AutoScaling"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743871) //= 0x6002003f. The runtime will prefer the VTID if present
  @VTID(70)
  boolean getAutoScaling();


  /**
   * <p>
   * Setter method for the COM property "AutoScaling"
   * </p>
   * @param f Mandatory boolean parameter.
   */

  @DISPID(1610743871) //= 0x6002003f. The runtime will prefer the VTID if present
  @VTID(71)
  void setAutoScaling(
    boolean f);


  /**
   * @param bstr Mandatory java.lang.String parameter.
   */

  @DISPID(1610743873) //= 0x60020041. The runtime will prefer the VTID if present
  @VTID(72)
  void setBackgroundPicture(
    java.lang.String bstr);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varSource is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varGallery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varPlotBy is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varCategoryLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varSeriesLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varHasLegend is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varCategoryTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varValueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varExtraTitle is set to com4j.Variant.getMissing()</li><li>int parameter localeID is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(1610743874) //= 0x60020042. The runtime will prefer the VTID if present
  @VTID(73)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void chartWizard();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varGallery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varPlotBy is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varCategoryLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varSeriesLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varHasLegend is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varCategoryTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varValueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varExtraTitle is set to com4j.Variant.getMissing()</li><li>int parameter localeID is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(varSource, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param varSource Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743874) //= 0x60020042. The runtime will prefer the VTID if present
  @VTID(73)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSource);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varPlotBy is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varCategoryLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varSeriesLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varHasLegend is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varCategoryTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varValueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varExtraTitle is set to com4j.Variant.getMissing()</li><li>int parameter localeID is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(varSource, varGallery, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param varSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varGallery Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743874) //= 0x60020042. The runtime will prefer the VTID if present
  @VTID(73)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSource,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varGallery);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varPlotBy is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varCategoryLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varSeriesLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varHasLegend is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varCategoryTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varValueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varExtraTitle is set to com4j.Variant.getMissing()</li><li>int parameter localeID is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(varSource, varGallery, varFormat, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param varSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varGallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varFormat Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743874) //= 0x60020042. The runtime will prefer the VTID if present
  @VTID(73)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSource,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varGallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varFormat);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varCategoryLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varSeriesLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varHasLegend is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varCategoryTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varValueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varExtraTitle is set to com4j.Variant.getMissing()</li><li>int parameter localeID is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(varSource, varGallery, varFormat, varPlotBy, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param varSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varGallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varPlotBy Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743874) //= 0x60020042. The runtime will prefer the VTID if present
  @VTID(73)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSource,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varGallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varPlotBy);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varSeriesLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varHasLegend is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varCategoryTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varValueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varExtraTitle is set to com4j.Variant.getMissing()</li><li>int parameter localeID is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param varSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varGallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varPlotBy Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varCategoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743874) //= 0x60020042. The runtime will prefer the VTID if present
  @VTID(73)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSource,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varGallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varPlotBy,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varCategoryLabels);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varHasLegend is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varCategoryTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varValueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varExtraTitle is set to com4j.Variant.getMissing()</li><li>int parameter localeID is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param varSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varGallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varPlotBy Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varCategoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varSeriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743874) //= 0x60020042. The runtime will prefer the VTID if present
  @VTID(73)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSource,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varGallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varPlotBy,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varCategoryLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSeriesLabels);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varCategoryTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varValueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varExtraTitle is set to com4j.Variant.getMissing()</li><li>int parameter localeID is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param varSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varGallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varPlotBy Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varCategoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varSeriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varHasLegend Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743874) //= 0x60020042. The runtime will prefer the VTID if present
  @VTID(73)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSource,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varGallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varPlotBy,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varCategoryLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSeriesLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varHasLegend);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varCategoryTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varValueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varExtraTitle is set to com4j.Variant.getMissing()</li><li>int parameter localeID is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend, varTitle, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param varSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varGallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varPlotBy Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varCategoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varSeriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varHasLegend Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varTitle Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743874) //= 0x60020042. The runtime will prefer the VTID if present
  @VTID(73)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSource,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varGallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varPlotBy,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varCategoryLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSeriesLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varHasLegend,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varTitle);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varValueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varExtraTitle is set to com4j.Variant.getMissing()</li><li>int parameter localeID is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend, varTitle, varCategoryTitle, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param varSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varGallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varPlotBy Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varCategoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varSeriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varHasLegend Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varTitle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varCategoryTitle Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743874) //= 0x60020042. The runtime will prefer the VTID if present
  @VTID(73)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSource,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varGallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varPlotBy,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varCategoryLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSeriesLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varHasLegend,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varTitle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varCategoryTitle);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varExtraTitle is set to com4j.Variant.getMissing()</li><li>int parameter localeID is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend, varTitle, varCategoryTitle, varValueTitle, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param varSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varGallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varPlotBy Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varCategoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varSeriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varHasLegend Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varTitle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varCategoryTitle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varValueTitle Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743874) //= 0x60020042. The runtime will prefer the VTID if present
  @VTID(73)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, optParamIndex = {10, 11}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSource,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varGallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varPlotBy,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varCategoryLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSeriesLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varHasLegend,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varTitle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varCategoryTitle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varValueTitle);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter localeID is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend, varTitle, varCategoryTitle, varValueTitle, varExtraTitle, 1033);
   * </code>
   * </p>
   * @param varSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varGallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varPlotBy Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varCategoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varSeriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varHasLegend Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varTitle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varCategoryTitle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varValueTitle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varExtraTitle Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743874) //= 0x60020042. The runtime will prefer the VTID if present
  @VTID(73)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, optParamIndex = {11}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSource,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varGallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varPlotBy,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varCategoryLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSeriesLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varHasLegend,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varTitle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varCategoryTitle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varValueTitle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varExtraTitle);

  /**
   * @param varSource Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varGallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varPlotBy Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varCategoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varSeriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varHasLegend Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varTitle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varCategoryTitle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varValueTitle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varExtraTitle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param localeID Mandatory int parameter.
   */

  @DISPID(1610743874) //= 0x60020042. The runtime will prefer the VTID if present
  @VTID(73)
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSource,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varGallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varPlotBy,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varCategoryLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varSeriesLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varHasLegend,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varTitle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varCategoryTitle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varValueTitle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varExtraTitle,
    @LCID int localeID);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter appearance is set to 1</li><li>int parameter format is set to -4147</li><li>int parameter size is set to 2</li><li>int parameter localeID is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copyPicture(1, -4147, 2, 1033);
   * </code>
   * </p>
   */

  @DISPID(1610743875) //= 0x60020043. The runtime will prefer the VTID if present
  @VTID(74)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {int.class, int.class, int.class, int.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "-4147", "2", "1033"})
  void copyPicture();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter format is set to -4147</li><li>int parameter size is set to 2</li><li>int parameter localeID is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copyPicture(appearance, -4147, 2, 1033);
   * </code>
   * </p>
   * @param appearance Optional parameter. Default value is 1
   */

  @DISPID(1610743875) //= 0x60020043. The runtime will prefer the VTID if present
  @VTID(74)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {int.class, int.class, int.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"-4147", "2", "1033"})
  void copyPicture(
    @Optional @DefaultValue("1") int appearance);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter size is set to 2</li><li>int parameter localeID is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copyPicture(appearance, format, 2, 1033);
   * </code>
   * </p>
   * @param appearance Optional parameter. Default value is 1
   * @param format Optional parameter. Default value is -4147
   */

  @DISPID(1610743875) //= 0x60020043. The runtime will prefer the VTID if present
  @VTID(74)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {int.class, int.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"2", "1033"})
  void copyPicture(
    @Optional @DefaultValue("1") int appearance,
    @Optional @DefaultValue("-4147") int format);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter localeID is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copyPicture(appearance, format, size, 1033);
   * </code>
   * </p>
   * @param appearance Optional parameter. Default value is 1
   * @param format Optional parameter. Default value is -4147
   * @param size Optional parameter. Default value is 2
   */

  @DISPID(1610743875) //= 0x60020043. The runtime will prefer the VTID if present
  @VTID(74)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void copyPicture(
    @Optional @DefaultValue("1") int appearance,
    @Optional @DefaultValue("-4147") int format,
    @Optional @DefaultValue("2") int size);

  /**
   * @param appearance Optional parameter. Default value is 1
   * @param format Optional parameter. Default value is -4147
   * @param size Optional parameter. Default value is 2
   * @param localeID Mandatory int parameter.
   */

  @DISPID(1610743875) //= 0x60020043. The runtime will prefer the VTID if present
  @VTID(74)
  void copyPicture(
    @Optional @DefaultValue("1") int appearance,
    @Optional @DefaultValue("-4147") int format,
    @Optional @DefaultValue("2") int size,
    @LCID int localeID);


  /**
   * <p>
   * Getter method for the COM property "DataTable"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoDataTable
   */

  @DISPID(1610743876) //= 0x60020044. The runtime will prefer the VTID if present
  @VTID(75)
  com.exceljava.com4j.office.IMsoDataTable getDataTable();


  /**
   * @param varName Mandatory java.lang.Object parameter.
   * @param localeID Mandatory int parameter.
   * @param objType Mandatory Holder<Integer> parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1610743877) //= 0x60020045. The runtime will prefer the VTID if present
  @VTID(76)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object evaluate(
    @MarshalAs(NativeType.VARIANT) java.lang.Object varName,
    int localeID,
    Holder<Integer> objType);


  /**
   * @param varName Mandatory java.lang.Object parameter.
   * @param localeID Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1610743878) //= 0x60020046. The runtime will prefer the VTID if present
  @VTID(77)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object _Evaluate(
    @MarshalAs(NativeType.VARIANT) java.lang.Object varName,
    int localeID);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varType is set to com4j.Variant.getMissing()</li><li>int parameter localeID is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * paste(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(1610743879) //= 0x60020047. The runtime will prefer the VTID if present
  @VTID(78)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void paste();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter localeID is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * paste(varType, 1033);
   * </code>
   * </p>
   * @param varType Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610743879) //= 0x60020047. The runtime will prefer the VTID if present
  @VTID(78)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void paste(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varType);

  /**
   * @param varType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param localeID Mandatory int parameter.
   */

  @DISPID(1610743879) //= 0x60020047. The runtime will prefer the VTID if present
  @VTID(78)
  void paste(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varType,
    @LCID int localeID);


  /**
   * <p>
   * Getter method for the COM property "BarShape"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.XlBarShape
   */

  @DISPID(1610743880) //= 0x60020048. The runtime will prefer the VTID if present
  @VTID(79)
  com.exceljava.com4j.office.XlBarShape getBarShape();


  /**
   * <p>
   * Setter method for the COM property "BarShape"
   * </p>
   * @param pShape Mandatory com.exceljava.com4j.office.XlBarShape parameter.
   */

  @DISPID(1610743880) //= 0x60020048. The runtime will prefer the VTID if present
  @VTID(80)
  void setBarShape(
    com.exceljava.com4j.office.XlBarShape pShape);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varFilterName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter varInteractive is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * export(bstr, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param bstr Mandatory java.lang.String parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743882) //= 0x6002004a. The runtime will prefer the VTID if present
  @VTID(81)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=3)
  boolean export(
    java.lang.String bstr);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varInteractive is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * export(bstr, varFilterName, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param bstr Mandatory java.lang.String parameter.
   * @param varFilterName Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743882) //= 0x6002004a. The runtime will prefer the VTID if present
  @VTID(81)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  boolean export(
    java.lang.String bstr,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varFilterName);

  /**
   * @param bstr Mandatory java.lang.String parameter.
   * @param varFilterName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param varInteractive Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743882) //= 0x6002004a. The runtime will prefer the VTID if present
  @VTID(81)
  boolean export(
    java.lang.String bstr,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varFilterName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varInteractive);


  /**
   * @param varName Mandatory java.lang.Object parameter.
   */

  @DISPID(1610743883) //= 0x6002004b. The runtime will prefer the VTID if present
  @VTID(82)
  void setDefaultChart(
    @MarshalAs(NativeType.VARIANT) java.lang.Object varName);


  /**
   * @param bstrFileName Mandatory java.lang.String parameter.
   */

  @DISPID(1610743884) //= 0x6002004c. The runtime will prefer the VTID if present
  @VTID(83)
  void applyChartTemplate(
    java.lang.String bstrFileName);


  /**
   * @param bstrFileName Mandatory java.lang.String parameter.
   */

  @DISPID(1610743885) //= 0x6002004d. The runtime will prefer the VTID if present
  @VTID(84)
  void saveChartTemplate(
    java.lang.String bstrFileName);


  /**
   * <p>
   * Getter method for the COM property "SideWall"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoWalls
   */

  @DISPID(2377) //= 0x949. The runtime will prefer the VTID if present
  @VTID(85)
  com.exceljava.com4j.office.IMsoWalls getSideWall();


  /**
   * <p>
   * Getter method for the COM property "BackWall"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoWalls
   */

  @DISPID(2378) //= 0x94a. The runtime will prefer the VTID if present
  @VTID(86)
  com.exceljava.com4j.office.IMsoWalls getBackWall();


  /**
   * <p>
   * Getter method for the COM property "ChartStyle"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(2465) //= 0x9a1. The runtime will prefer the VTID if present
  @VTID(87)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getChartStyle();


  /**
   * <p>
   * Setter method for the COM property "ChartStyle"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(2465) //= 0x9a1. The runtime will prefer the VTID if present
  @VTID(88)
  void setChartStyle(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   */

  @DISPID(2466) //= 0x9a2. The runtime will prefer the VTID if present
  @VTID(89)
  void clearToMatchStyle();


  /**
   * <p>
   * Getter method for the COM property "PivotLayout"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1814) //= 0x716. The runtime will prefer the VTID if present
  @VTID(90)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getPivotLayout();


  /**
   * <p>
   * Getter method for the COM property "HasPivotFields"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1815) //= 0x717. The runtime will prefer the VTID if present
  @VTID(91)
  boolean getHasPivotFields();


  /**
   * <p>
   * Setter method for the COM property "HasPivotFields"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1815) //= 0x717. The runtime will prefer the VTID if present
  @VTID(92)
  void setHasPivotFields(
    boolean rhs);


  /**
   */

  @DISPID(1610743894) //= 0x60020056. The runtime will prefer the VTID if present
  @VTID(93)
  void refreshPivotTable();


  /**
   * <p>
   * Setter method for the COM property "ShowDataLabelsOverMaximum"
   * </p>
   * @param pRHS Mandatory boolean parameter.
   */

  @DISPID(1610743895) //= 0x60020057. The runtime will prefer the VTID if present
  @VTID(94)
  void setShowDataLabelsOverMaximum(
    boolean pRHS);


  /**
   * <p>
   * Getter method for the COM property "ShowDataLabelsOverMaximum"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743895) //= 0x60020057. The runtime will prefer the VTID if present
  @VTID(95)
  boolean getShowDataLabelsOverMaximum();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varChartType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyLayout(layout, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param layout Mandatory int parameter.
   */

  @DISPID(2468) //= 0x9a4. The runtime will prefer the VTID if present
  @VTID(96)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void applyLayout(
    int layout);

  /**
   * @param layout Mandatory int parameter.
   * @param varChartType Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2468) //= 0x9a4. The runtime will prefer the VTID if present
  @VTID(96)
  void applyLayout(
    int layout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varChartType);


  /**
   * <p>
   * Getter method for the COM property "Selection"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610743898) //= 0x6002005a. The runtime will prefer the VTID if present
  @VTID(97)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getSelection();


  /**
   */

  @DISPID(1610743899) //= 0x6002005b. The runtime will prefer the VTID if present
  @VTID(98)
  void refresh();


  /**
   * @param rhs Mandatory com.exceljava.com4j.office.MsoChartElementType parameter.
   */

  @DISPID(1610743900) //= 0x6002005c. The runtime will prefer the VTID if present
  @VTID(99)
  void setElement(
    com.exceljava.com4j.office.MsoChartElementType rhs);


  /**
   * <p>
   * Getter method for the COM property "ChartData"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChartData
   */

  @DISPID(1610743901) //= 0x6002005d. The runtime will prefer the VTID if present
  @VTID(100)
  com.exceljava.com4j.office.IMsoChartData getChartData();


  /**
   * <p>
   * Getter method for the COM property "Format"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChartFormat
   */

  @DISPID(1610743902) //= 0x6002005e. The runtime will prefer the VTID if present
  @VTID(101)
  com.exceljava.com4j.office.IMsoChartFormat getFormat();


  /**
   * <p>
   * Getter method for the COM property "Shapes"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Shapes
   */

  @DISPID(1610743903) //= 0x6002005f. The runtime will prefer the VTID if present
  @VTID(102)
  com.exceljava.com4j.office.Shapes getShapes();


  @VTID(102)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.Shapes.class})
  com.exceljava.com4j.office.Shape getShapes(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(148) //= 0x94. The runtime will prefer the VTID if present
  @VTID(103)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(149) //= 0x95. The runtime will prefer the VTID if present
  @VTID(104)
  int getCreator();


  /**
   * <p>
   * Getter method for the COM property "Area3DGroup"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getArea3DGroup(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChartGroup
   */

  @DISPID(17) //= 0x11. The runtime will prefer the VTID if present
  @VTID(105)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.office.IMsoChartGroup getArea3DGroup();

  /**
   * <p>
   * Getter method for the COM property "Area3DGroup"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChartGroup
   */

  @DISPID(17) //= 0x11. The runtime will prefer the VTID if present
  @VTID(105)
  com.exceljava.com4j.office.IMsoChartGroup getArea3DGroup(
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * areaGroups(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(106)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject areaGroups();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * areaGroups(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(106)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject areaGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(106)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject areaGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Bar3DGroup"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getBar3DGroup(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChartGroup
   */

  @DISPID(18) //= 0x12. The runtime will prefer the VTID if present
  @VTID(107)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.office.IMsoChartGroup getBar3DGroup();

  /**
   * <p>
   * Getter method for the COM property "Bar3DGroup"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChartGroup
   */

  @DISPID(18) //= 0x12. The runtime will prefer the VTID if present
  @VTID(107)
  com.exceljava.com4j.office.IMsoChartGroup getBar3DGroup(
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * barGroups(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(108)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject barGroups();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * barGroups(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(108)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject barGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(108)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject barGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Column3DGroup"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getColumn3DGroup(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChartGroup
   */

  @DISPID(19) //= 0x13. The runtime will prefer the VTID if present
  @VTID(109)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.office.IMsoChartGroup getColumn3DGroup();

  /**
   * <p>
   * Getter method for the COM property "Column3DGroup"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChartGroup
   */

  @DISPID(19) //= 0x13. The runtime will prefer the VTID if present
  @VTID(109)
  com.exceljava.com4j.office.IMsoChartGroup getColumn3DGroup(
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * columnGroups(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(110)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject columnGroups();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * columnGroups(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(110)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject columnGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(110)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject columnGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Line3DGroup"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getLine3DGroup(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChartGroup
   */

  @DISPID(20) //= 0x14. The runtime will prefer the VTID if present
  @VTID(111)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.office.IMsoChartGroup getLine3DGroup();

  /**
   * <p>
   * Getter method for the COM property "Line3DGroup"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChartGroup
   */

  @DISPID(20) //= 0x14. The runtime will prefer the VTID if present
  @VTID(111)
  com.exceljava.com4j.office.IMsoChartGroup getLine3DGroup(
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * lineGroups(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(112)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject lineGroups();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * lineGroups(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(112)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject lineGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(112)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject lineGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Pie3DGroup"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getPie3DGroup(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChartGroup
   */

  @DISPID(21) //= 0x15. The runtime will prefer the VTID if present
  @VTID(113)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.office.IMsoChartGroup getPie3DGroup();

  /**
   * <p>
   * Getter method for the COM property "Pie3DGroup"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChartGroup
   */

  @DISPID(21) //= 0x15. The runtime will prefer the VTID if present
  @VTID(113)
  com.exceljava.com4j.office.IMsoChartGroup getPie3DGroup(
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pieGroups(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(114)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject pieGroups();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pieGroups(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(114)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject pieGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(114)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject pieGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * doughnutGroups(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(115)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject doughnutGroups();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * doughnutGroups(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(115)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject doughnutGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(115)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject doughnutGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * radarGroups(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(116)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject radarGroups();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * radarGroups(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(116)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject radarGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(116)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject radarGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "SurfaceGroup"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getSurfaceGroup(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChartGroup
   */

  @DISPID(22) //= 0x16. The runtime will prefer the VTID if present
  @VTID(117)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.office.IMsoChartGroup getSurfaceGroup();

  /**
   * <p>
   * Getter method for the COM property "SurfaceGroup"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoChartGroup
   */

  @DISPID(22) //= 0x16. The runtime will prefer the VTID if present
  @VTID(117)
  com.exceljava.com4j.office.IMsoChartGroup getSurfaceGroup(
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * xyGroups(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(16) //= 0x10. The runtime will prefer the VTID if present
  @VTID(118)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject xyGroups();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * xyGroups(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(16) //= 0x10. The runtime will prefer the VTID if present
  @VTID(118)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject xyGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(16) //= 0x10. The runtime will prefer the VTID if present
  @VTID(118)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject xyGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(117) //= 0x75. The runtime will prefer the VTID if present
  @VTID(119)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object delete();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(551) //= 0x227. The runtime will prefer the VTID if present
  @VTID(120)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object copy();


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
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(235) //= 0xeb. The runtime will prefer the VTID if present
  @VTID(121)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object select();

  /**
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(235) //= 0xeb. The runtime will prefer the VTID if present
  @VTID(121)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object select(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object replace);


  /**
   * <p>
   * Getter method for the COM property "ShowReportFilterFieldButtons"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743923) //= 0x60020073. The runtime will prefer the VTID if present
  @VTID(122)
  boolean getShowReportFilterFieldButtons();


  /**
   * <p>
   * Setter method for the COM property "ShowReportFilterFieldButtons"
   * </p>
   * @param res Mandatory boolean parameter.
   */

  @DISPID(1610743923) //= 0x60020073. The runtime will prefer the VTID if present
  @VTID(123)
  void setShowReportFilterFieldButtons(
    boolean res);


  /**
   * <p>
   * Getter method for the COM property "ShowLegendFieldButtons"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743925) //= 0x60020075. The runtime will prefer the VTID if present
  @VTID(124)
  boolean getShowLegendFieldButtons();


  /**
   * <p>
   * Setter method for the COM property "ShowLegendFieldButtons"
   * </p>
   * @param res Mandatory boolean parameter.
   */

  @DISPID(1610743925) //= 0x60020075. The runtime will prefer the VTID if present
  @VTID(125)
  void setShowLegendFieldButtons(
    boolean res);


  /**
   * <p>
   * Getter method for the COM property "ShowAxisFieldButtons"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743927) //= 0x60020077. The runtime will prefer the VTID if present
  @VTID(126)
  boolean getShowAxisFieldButtons();


  /**
   * <p>
   * Setter method for the COM property "ShowAxisFieldButtons"
   * </p>
   * @param res Mandatory boolean parameter.
   */

  @DISPID(1610743927) //= 0x60020077. The runtime will prefer the VTID if present
  @VTID(127)
  void setShowAxisFieldButtons(
    boolean res);


  /**
   * <p>
   * Getter method for the COM property "ShowValueFieldButtons"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743929) //= 0x60020079. The runtime will prefer the VTID if present
  @VTID(128)
  boolean getShowValueFieldButtons();


  /**
   * <p>
   * Setter method for the COM property "ShowValueFieldButtons"
   * </p>
   * @param res Mandatory boolean parameter.
   */

  @DISPID(1610743929) //= 0x60020079. The runtime will prefer the VTID if present
  @VTID(129)
  void setShowValueFieldButtons(
    boolean res);


  /**
   * <p>
   * Getter method for the COM property "ShowAllFieldButtons"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743931) //= 0x6002007b. The runtime will prefer the VTID if present
  @VTID(130)
  boolean getShowAllFieldButtons();


  /**
   * <p>
   * Setter method for the COM property "ShowAllFieldButtons"
   * </p>
   * @param res Mandatory boolean parameter.
   */

  @DISPID(1610743931) //= 0x6002007b. The runtime will prefer the VTID if present
  @VTID(131)
  void setShowAllFieldButtons(
    boolean res);


  /**
   * <p>
   * Setter method for the COM property "ProtectChartSheetFormatting"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1610743933) //= 0x6002007d. The runtime will prefer the VTID if present
  @VTID(132)
  void setProtectChartSheetFormatting(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * fullSeriesCollection(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(236) //= 0xec. The runtime will prefer the VTID if present
  @VTID(133)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.Dispatch,index=1)
  com4j.Com4jObject fullSeriesCollection();

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(236) //= 0xec. The runtime will prefer the VTID if present
  @VTID(133)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject fullSeriesCollection(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Setter method for the COM property "CategoryLabelLevel"
   * </p>
   * @param plevel Mandatory com.exceljava.com4j.office.XlCategoryLabelLevel parameter.
   */

  @DISPID(237) //= 0xed. The runtime will prefer the VTID if present
  @VTID(134)
  void setCategoryLabelLevel(
    com.exceljava.com4j.office.XlCategoryLabelLevel plevel);


  /**
   * <p>
   * Getter method for the COM property "CategoryLabelLevel"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.XlCategoryLabelLevel
   */

  @DISPID(237) //= 0xed. The runtime will prefer the VTID if present
  @VTID(135)
  com.exceljava.com4j.office.XlCategoryLabelLevel getCategoryLabelLevel();


  /**
   * <p>
   * Setter method for the COM property "SeriesNameLevel"
   * </p>
   * @param plevel Mandatory com.exceljava.com4j.office.XlSeriesNameLevel parameter.
   */

  @DISPID(238) //= 0xee. The runtime will prefer the VTID if present
  @VTID(136)
  void setSeriesNameLevel(
    com.exceljava.com4j.office.XlSeriesNameLevel plevel);


  /**
   * <p>
   * Getter method for the COM property "SeriesNameLevel"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.XlSeriesNameLevel
   */

  @DISPID(238) //= 0xee. The runtime will prefer the VTID if present
  @VTID(137)
  com.exceljava.com4j.office.XlSeriesNameLevel getSeriesNameLevel();


  /**
   * <p>
   * Getter method for the COM property "HasHiddenContent"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(239) //= 0xef. The runtime will prefer the VTID if present
  @VTID(138)
  boolean getHasHiddenContent();


  /**
   */

  @DISPID(240) //= 0xf0. The runtime will prefer the VTID if present
  @VTID(139)
  void deleteHiddenContent();


  /**
   * <p>
   * Getter method for the COM property "ChartColor"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(2467) //= 0x9a3. The runtime will prefer the VTID if present
  @VTID(140)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getChartColor();


  /**
   * <p>
   * Setter method for the COM property "ChartColor"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(2467) //= 0x9a3. The runtime will prefer the VTID if present
  @VTID(141)
  void setChartColor(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   */

  @DISPID(2469) //= 0x9a5. The runtime will prefer the VTID if present
  @VTID(142)
  void clearToMatchColorStyle();


  /**
   * <p>
   * Getter method for the COM property "ShowExpandCollapseEntireFieldButtons"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743944) //= 0x60020088. The runtime will prefer the VTID if present
  @VTID(143)
  boolean getShowExpandCollapseEntireFieldButtons();


  /**
   * <p>
   * Setter method for the COM property "ShowExpandCollapseEntireFieldButtons"
   * </p>
   * @param res Mandatory boolean parameter.
   */

  @DISPID(1610743944) //= 0x60020088. The runtime will prefer the VTID if present
  @VTID(144)
  void setShowExpandCollapseEntireFieldButtons(
    boolean res);


  /**
   * @param bstrId Mandatory java.lang.String parameter.
   * @param value Mandatory java.lang.Object parameter.
   */

  @DISPID(253) //= 0xfd. The runtime will prefer the VTID if present
  @VTID(145)
  void setProperty(
    java.lang.String bstrId,
    @MarshalAs(NativeType.VARIANT) java.lang.Object value);


  /**
   * @param bstrId Mandatory java.lang.String parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(254) //= 0xfe. The runtime will prefer the VTID if present
  @VTID(146)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getProperty(
    java.lang.String bstrId);


  /**
   * <p>
   * Getter method for the COM property "DisplayValueNotAvailableAsBlank"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610743948) //= 0x6002008c. The runtime will prefer the VTID if present
  @VTID(147)
  boolean getDisplayValueNotAvailableAsBlank();


  /**
   * <p>
   * Setter method for the COM property "DisplayValueNotAvailableAsBlank"
   * </p>
   * @param res Mandatory boolean parameter.
   */

  @DISPID(1610743948) //= 0x6002008c. The runtime will prefer the VTID if present
  @VTID(148)
  void setDisplayValueNotAvailableAsBlank(
    boolean res);


  // Properties:
}
