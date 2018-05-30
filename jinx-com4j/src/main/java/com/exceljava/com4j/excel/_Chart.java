package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000208D6-0000-0000-C000-000000000046}")
public interface _Chart extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Application
   */

  @DISPID(148) //= 0x94. The runtime will prefer the VTID if present
  @VTID(7)
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCreator
   */

  @DISPID(149) //= 0x95. The runtime will prefer the VTID if present
  @VTID(8)
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(150) //= 0x96. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * activate(1033);
   * </code>
   * </p>
   */

  @DISPID(304) //= 0x130. The runtime will prefer the VTID if present
  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void activate();

  /**
   * @param lcid Mandatory int parameter.
   */

  @DISPID(304) //= 0x130. The runtime will prefer the VTID if present
  @VTID(10)
  void activate(
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter before is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter after is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copy(com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(551) //= 0x227. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void copy();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter after is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copy(before, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(551) //= 0x227. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void copy(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copy(before, after, 1033);
   * </code>
   * </p>
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @param after Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(551) //= 0x227. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void copy(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object after);

  /**
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @param after Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(551) //= 0x227. The runtime will prefer the VTID if present
  @VTID(11)
  void copy(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object after,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * delete(1033);
   * </code>
   * </p>
   */

  @DISPID(117) //= 0x75. The runtime will prefer the VTID if present
  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void delete();

  /**
   * @param lcid Mandatory int parameter.
   */

  @DISPID(117) //= 0x75. The runtime will prefer the VTID if present
  @VTID(12)
  void delete(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "CodeName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1373) //= 0x55d. The runtime will prefer the VTID if present
  @VTID(13)
  java.lang.String getCodeName();


  /**
   * <p>
   * Getter method for the COM property "_CodeName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(-2147418112) //= 0x80010000. The runtime will prefer the VTID if present
  @VTID(14)
  java.lang.String get_CodeName();


  /**
   * <p>
   * Setter method for the COM property "_CodeName"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(-2147418112) //= 0x80010000. The runtime will prefer the VTID if present
  @VTID(15)
  void set_CodeName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Index"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getIndex(1033);
   * </code>
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(486) //= 0x1e6. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  int getIndex();

  /**
   * <p>
   * Getter method for the COM property "Index"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type int
   */

  @DISPID(486) //= 0x1e6. The runtime will prefer the VTID if present
  @VTID(16)
  int getIndex(
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter before is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter after is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * move(com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(637) //= 0x27d. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void move();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter after is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * move(before, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(637) //= 0x27d. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void move(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * move(before, after, 1033);
   * </code>
   * </p>
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @param after Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(637) //= 0x27d. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void move(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object after);

  /**
   * @param before Optional parameter. Default value is com4j.Variant.getMissing()
   * @param after Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(637) //= 0x27d. The runtime will prefer the VTID if present
  @VTID(17)
  void move(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object before,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object after,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(110) //= 0x6e. The runtime will prefer the VTID if present
  @VTID(18)
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(110) //= 0x6e. The runtime will prefer the VTID if present
  @VTID(19)
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Next"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(502) //= 0x1f6. The runtime will prefer the VTID if present
  @VTID(20)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getNext();


  /**
   * <p>
   * Getter method for the COM property "OnDoubleClick"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getOnDoubleClick(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(628) //= 0x274. The runtime will prefer the VTID if present
  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getOnDoubleClick();

  /**
   * <p>
   * Getter method for the COM property "OnDoubleClick"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(628) //= 0x274. The runtime will prefer the VTID if present
  @VTID(21)
  java.lang.String getOnDoubleClick(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "OnDoubleClick"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setOnDoubleClick(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(628) //= 0x274. The runtime will prefer the VTID if present
  @VTID(22)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setOnDoubleClick(
    java.lang.String rhs);

  /**
   * <p>
   * Setter method for the COM property "OnDoubleClick"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(628) //= 0x274. The runtime will prefer the VTID if present
  @VTID(22)
  void setOnDoubleClick(
    @LCID int lcid,
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "OnSheetActivate"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getOnSheetActivate(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1031) //= 0x407. The runtime will prefer the VTID if present
  @VTID(23)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getOnSheetActivate();

  /**
   * <p>
   * Getter method for the COM property "OnSheetActivate"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1031) //= 0x407. The runtime will prefer the VTID if present
  @VTID(23)
  java.lang.String getOnSheetActivate(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "OnSheetActivate"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setOnSheetActivate(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1031) //= 0x407. The runtime will prefer the VTID if present
  @VTID(24)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setOnSheetActivate(
    java.lang.String rhs);

  /**
   * <p>
   * Setter method for the COM property "OnSheetActivate"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1031) //= 0x407. The runtime will prefer the VTID if present
  @VTID(24)
  void setOnSheetActivate(
    @LCID int lcid,
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "OnSheetDeactivate"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getOnSheetDeactivate(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1081) //= 0x439. The runtime will prefer the VTID if present
  @VTID(25)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getOnSheetDeactivate();

  /**
   * <p>
   * Getter method for the COM property "OnSheetDeactivate"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1081) //= 0x439. The runtime will prefer the VTID if present
  @VTID(25)
  java.lang.String getOnSheetDeactivate(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "OnSheetDeactivate"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setOnSheetDeactivate(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1081) //= 0x439. The runtime will prefer the VTID if present
  @VTID(26)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setOnSheetDeactivate(
    java.lang.String rhs);

  /**
   * <p>
   * Setter method for the COM property "OnSheetDeactivate"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1081) //= 0x439. The runtime will prefer the VTID if present
  @VTID(26)
  void setOnSheetDeactivate(
    @LCID int lcid,
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "PageSetup"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PageSetup
   */

  @DISPID(998) //= 0x3e6. The runtime will prefer the VTID if present
  @VTID(27)
  com.exceljava.com4j.excel.PageSetup getPageSetup();


  /**
   * <p>
   * Getter method for the COM property "Previous"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(503) //= 0x1f7. The runtime will prefer the VTID if present
  @VTID(28)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getPrevious();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(905) //= 0x389. The runtime will prefer the VTID if present
  @VTID(29)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __PrintOut();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905) //= 0x389. The runtime will prefer the VTID if present
  @VTID(29)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, to, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905) //= 0x389. The runtime will prefer the VTID if present
  @VTID(29)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, to, copies, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905) //= 0x389. The runtime will prefer the VTID if present
  @VTID(29)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, to, copies, preview, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905) //= 0x389. The runtime will prefer the VTID if present
  @VTID(29)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, to, copies, preview, activePrinter, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905) //= 0x389. The runtime will prefer the VTID if present
  @VTID(29)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, to, copies, preview, activePrinter, printToFile, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905) //= 0x389. The runtime will prefer the VTID if present
  @VTID(29)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __PrintOut(from, to, copies, preview, activePrinter, printToFile, collate, 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(905) //= 0x389. The runtime will prefer the VTID if present
  @VTID(29)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate);

  /**
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(905) //= 0x389. The runtime will prefer the VTID if present
  @VTID(29)
  void __PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter enableChanges is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printPreview(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(281) //= 0x119. The runtime will prefer the VTID if present
  @VTID(30)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void printPreview();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printPreview(enableChanges, 1033);
   * </code>
   * </p>
   * @param enableChanges Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(281) //= 0x119. The runtime will prefer the VTID if present
  @VTID(30)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void printPreview(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object enableChanges);

  /**
   * @param enableChanges Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(281) //= 0x119. The runtime will prefer the VTID if present
  @VTID(30)
  void printPreview(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object enableChanges,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter drawingObjects is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter contents is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter scenarios is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter userInterfaceOnly is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Protect(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(282) //= 0x11a. The runtime will prefer the VTID if present
  @VTID(31)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _Protect();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter drawingObjects is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter contents is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter scenarios is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter userInterfaceOnly is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Protect(password, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(282) //= 0x11a. The runtime will prefer the VTID if present
  @VTID(31)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void _Protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter contents is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter scenarios is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter userInterfaceOnly is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Protect(password, drawingObjects, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(282) //= 0x11a. The runtime will prefer the VTID if present
  @VTID(31)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void _Protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter scenarios is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter userInterfaceOnly is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Protect(password, drawingObjects, contents, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(282) //= 0x11a. The runtime will prefer the VTID if present
  @VTID(31)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void _Protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter userInterfaceOnly is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Protect(password, drawingObjects, contents, scenarios, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scenarios Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(282) //= 0x11a. The runtime will prefer the VTID if present
  @VTID(31)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void _Protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scenarios);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Protect(password, drawingObjects, contents, scenarios, userInterfaceOnly, 1033);
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scenarios Optional parameter. Default value is com4j.Variant.getMissing()
   * @param userInterfaceOnly Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(282) //= 0x11a. The runtime will prefer the VTID if present
  @VTID(31)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void _Protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scenarios,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object userInterfaceOnly);

  /**
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scenarios Optional parameter. Default value is com4j.Variant.getMissing()
   * @param userInterfaceOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(282) //= 0x11a. The runtime will prefer the VTID if present
  @VTID(31)
  void _Protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scenarios,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object userInterfaceOnly,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "ProtectContents"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getProtectContents(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(292) //= 0x124. The runtime will prefer the VTID if present
  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getProtectContents();

  /**
   * <p>
   * Getter method for the COM property "ProtectContents"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(292) //= 0x124. The runtime will prefer the VTID if present
  @VTID(32)
  boolean getProtectContents(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "ProtectDrawingObjects"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getProtectDrawingObjects(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(293) //= 0x125. The runtime will prefer the VTID if present
  @VTID(33)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getProtectDrawingObjects();

  /**
   * <p>
   * Getter method for the COM property "ProtectDrawingObjects"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(293) //= 0x125. The runtime will prefer the VTID if present
  @VTID(33)
  boolean getProtectDrawingObjects(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "ProtectionMode"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getProtectionMode(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1159) //= 0x487. The runtime will prefer the VTID if present
  @VTID(34)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getProtectionMode();

  /**
   * <p>
   * Getter method for the COM property "ProtectionMode"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(1159) //= 0x487. The runtime will prefer the VTID if present
  @VTID(34)
  boolean getProtectionMode(
    @LCID int lcid);


  /**
   */

  @DISPID(65559) //= 0x10017. The runtime will prefer the VTID if present
  @VTID(35)
  void _Dummy23();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter fileFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(36)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __SaveAs(
    java.lang.String filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, fileFormat, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(36)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __SaveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, fileFormat, password, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(36)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __SaveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, fileFormat, password, writeResPassword, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(36)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void __SaveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(36)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void __SaveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(36)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void __SaveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(36)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void __SaveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(36)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void __SaveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * __SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage, textVisualLayout, 1033);
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(36)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void __SaveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout);

  /**
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(284) //= 0x11c. The runtime will prefer the VTID if present
  @VTID(36)
  void __SaveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter replace is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * select(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(235) //= 0xeb. The runtime will prefer the VTID if present
  @VTID(37)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void select();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * select(replace, 1033);
   * </code>
   * </p>
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(235) //= 0xeb. The runtime will prefer the VTID if present
  @VTID(37)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void select(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object replace);

  /**
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(235) //= 0xeb. The runtime will prefer the VTID if present
  @VTID(37)
  void select(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object replace,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * unprotect(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(285) //= 0x11d. The runtime will prefer the VTID if present
  @VTID(38)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void unprotect();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * unprotect(password, 1033);
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(285) //= 0x11d. The runtime will prefer the VTID if present
  @VTID(38)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void unprotect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);

  /**
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(285) //= 0x11d. The runtime will prefer the VTID if present
  @VTID(38)
  void unprotect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getVisible(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSheetVisibility
   */

  @DISPID(558) //= 0x22e. The runtime will prefer the VTID if present
  @VTID(39)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.XlSheetVisibility getVisible();

  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSheetVisibility
   */

  @DISPID(558) //= 0x22e. The runtime will prefer the VTID if present
  @VTID(39)
  com.exceljava.com4j.excel.XlSheetVisibility getVisible(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setVisible(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSheetVisibility parameter.
   */

  @DISPID(558) //= 0x22e. The runtime will prefer the VTID if present
  @VTID(40)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setVisible(
    com.exceljava.com4j.excel.XlSheetVisibility rhs);

  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSheetVisibility parameter.
   */

  @DISPID(558) //= 0x22e. The runtime will prefer the VTID if present
  @VTID(40)
  void setVisible(
    @LCID int lcid,
    com.exceljava.com4j.excel.XlSheetVisibility rhs);


  /**
   * <p>
   * Getter method for the COM property "Shapes"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Shapes
   */

  @DISPID(1377) //= 0x561. The runtime will prefer the VTID if present
  @VTID(41)
  com.exceljava.com4j.excel.Shapes getShapes();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlDataLabelsType parameter type is set to 2</li><li>java.lang.Object parameter legendKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ApplyDataLabels(2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(151) //= 0x97. The runtime will prefer the VTID if present
  @VTID(42)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {com.exceljava.com4j.excel.XlDataLabelsType.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"2", "80020004", "80020004", "80020004", "1033"})
  void _ApplyDataLabels();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter legendKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ApplyDataLabels(type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   */

  @DISPID(151) //= 0x97. The runtime will prefer the VTID if present
  @VTID(42)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void _ApplyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter autoText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ApplyDataLabels(type, legendKey, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(151) //= 0x97. The runtime will prefer the VTID if present
  @VTID(42)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void _ApplyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ApplyDataLabels(type, legendKey, autoText, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(151) //= 0x97. The runtime will prefer the VTID if present
  @VTID(42)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void _ApplyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ApplyDataLabels(type, legendKey, autoText, hasLeaderLines, 1033);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(151) //= 0x97. The runtime will prefer the VTID if present
  @VTID(42)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void _ApplyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines);

  /**
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(151) //= 0x97. The runtime will prefer the VTID if present
  @VTID(42)
  void _ApplyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines,
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
   * arcs(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(760) //= 0x2f8. The runtime will prefer the VTID if present
  @VTID(43)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject arcs();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * arcs(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(760) //= 0x2f8. The runtime will prefer the VTID if present
  @VTID(43)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject arcs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(760) //= 0x2f8. The runtime will prefer the VTID if present
  @VTID(43)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject arcs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


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
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartGroup
   */

  @DISPID(17) //= 0x11. The runtime will prefer the VTID if present
  @VTID(44)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.ChartGroup getArea3DGroup();

  /**
   * <p>
   * Getter method for the COM property "Area3DGroup"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartGroup
   */

  @DISPID(17) //= 0x11. The runtime will prefer the VTID if present
  @VTID(44)
  com.exceljava.com4j.excel.ChartGroup getArea3DGroup(
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
  @VTID(45)
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
  @VTID(45)
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
  @VTID(45)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject areaGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * autoFormat(gallery, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param gallery Mandatory int parameter.
   */

  @DISPID(114) //= 0x72. The runtime will prefer the VTID if present
  @VTID(46)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void autoFormat(
    int gallery);

  /**
   * @param gallery Mandatory int parameter.
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(114) //= 0x72. The runtime will prefer the VTID if present
  @VTID(46)
  void autoFormat(
    int gallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format);


  /**
   * <p>
   * Getter method for the COM property "AutoScaling"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAutoScaling(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getAutoScaling();

  /**
   * <p>
   * Getter method for the COM property "AutoScaling"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(47)
  boolean getAutoScaling(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "AutoScaling"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setAutoScaling(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(48)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setAutoScaling(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "AutoScaling"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(48)
  void setAutoScaling(
    @LCID int lcid,
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlAxisGroup parameter axisGroup is set to 1</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * axes(com4j.Variant.getMissing(), 1, 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(23) //= 0x17. The runtime will prefer the VTID if present
  @VTID(49)
  @UseDefaultValues(paramIndexMapping = {3}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlAxisGroup.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"80020004", "1", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=3)
  com4j.Com4jObject axes();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlAxisGroup parameter axisGroup is set to 1</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * axes(type, 1, 1033);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(23) //= 0x17. The runtime will prefer the VTID if present
  @VTID(49)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {com.exceljava.com4j.excel.XlAxisGroup.class, int.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=3)
  com4j.Com4jObject axes(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * axes(type, axisGroup, 1033);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param axisGroup Optional parameter. Default value is 1
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(23) //= 0x17. The runtime will prefer the VTID if present
  @VTID(49)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=3)
  com4j.Com4jObject axes(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlAxisGroup axisGroup);

  /**
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param axisGroup Optional parameter. Default value is 1
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(23) //= 0x17. The runtime will prefer the VTID if present
  @VTID(49)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject axes(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlAxisGroup axisGroup,
    @LCID int lcid);


  /**
   * @param filename Mandatory java.lang.String parameter.
   */

  @DISPID(1188) //= 0x4a4. The runtime will prefer the VTID if present
  @VTID(50)
  void setBackgroundPicture(
    java.lang.String filename);


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
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartGroup
   */

  @DISPID(18) //= 0x12. The runtime will prefer the VTID if present
  @VTID(51)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.ChartGroup getBar3DGroup();

  /**
   * <p>
   * Getter method for the COM property "Bar3DGroup"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartGroup
   */

  @DISPID(18) //= 0x12. The runtime will prefer the VTID if present
  @VTID(51)
  com.exceljava.com4j.excel.ChartGroup getBar3DGroup(
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
  @VTID(52)
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
  @VTID(52)
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
  @VTID(52)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject barGroups(
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
   * buttons(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(557) //= 0x22d. The runtime will prefer the VTID if present
  @VTID(53)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject buttons();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * buttons(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(557) //= 0x22d. The runtime will prefer the VTID if present
  @VTID(53)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject buttons(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(557) //= 0x22d. The runtime will prefer the VTID if present
  @VTID(53)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject buttons(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "ChartArea"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getChartArea(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartArea
   */

  @DISPID(80) //= 0x50. The runtime will prefer the VTID if present
  @VTID(54)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.ChartArea getChartArea();

  /**
   * <p>
   * Getter method for the COM property "ChartArea"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartArea
   */

  @DISPID(80) //= 0x50. The runtime will prefer the VTID if present
  @VTID(54)
  com.exceljava.com4j.excel.ChartArea getChartArea(
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
   * chartGroups(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(55)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject chartGroups();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartGroups(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(55)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject chartGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(55)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject chartGroups(
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
   * chartObjects(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1060) //= 0x424. The runtime will prefer the VTID if present
  @VTID(56)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject chartObjects();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartObjects(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1060) //= 0x424. The runtime will prefer the VTID if present
  @VTID(56)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject chartObjects(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1060) //= 0x424. The runtime will prefer the VTID if present
  @VTID(56)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject chartObjects(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "ChartTitle"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getChartTitle(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartTitle
   */

  @DISPID(81) //= 0x51. The runtime will prefer the VTID if present
  @VTID(57)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.ChartTitle getChartTitle();

  /**
   * <p>
   * Getter method for the COM property "ChartTitle"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartTitle
   */

  @DISPID(81) //= 0x51. The runtime will prefer the VTID if present
  @VTID(57)
  com.exceljava.com4j.excel.ChartTitle getChartTitle(
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter source is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter gallery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter plotBy is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter seriesLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLegend is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter title is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter valueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter extraTitle is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(196) //= 0xc4. The runtime will prefer the VTID if present
  @VTID(58)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void chartWizard();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter gallery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter plotBy is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter seriesLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLegend is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter title is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter valueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter extraTitle is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(source, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(196) //= 0xc4. The runtime will prefer the VTID if present
  @VTID(58)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter plotBy is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter seriesLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLegend is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter title is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter valueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter extraTitle is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(source, gallery, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param gallery Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(196) //= 0xc4. The runtime will prefer the VTID if present
  @VTID(58)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object gallery);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter plotBy is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter seriesLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLegend is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter title is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter valueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter extraTitle is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(source, gallery, format, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param gallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(196) //= 0xc4. The runtime will prefer the VTID if present
  @VTID(58)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object gallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter categoryLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter seriesLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLegend is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter title is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter valueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter extraTitle is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(source, gallery, format, plotBy, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param gallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param plotBy Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(196) //= 0xc4. The runtime will prefer the VTID if present
  @VTID(58)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object gallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object plotBy);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter seriesLabels is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLegend is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter title is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter valueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter extraTitle is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(source, gallery, format, plotBy, categoryLabels, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param gallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param plotBy Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(196) //= 0xc4. The runtime will prefer the VTID if present
  @VTID(58)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object gallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object plotBy,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryLabels);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter hasLegend is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter title is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter valueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter extraTitle is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(source, gallery, format, plotBy, categoryLabels, seriesLabels, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param gallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param plotBy Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(196) //= 0xc4. The runtime will prefer the VTID if present
  @VTID(58)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object gallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object plotBy,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seriesLabels);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter title is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter categoryTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter valueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter extraTitle is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param gallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param plotBy Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLegend Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(196) //= 0xc4. The runtime will prefer the VTID if present
  @VTID(58)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object gallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object plotBy,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seriesLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLegend);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter categoryTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter valueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter extraTitle is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param gallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param plotBy Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLegend Optional parameter. Default value is com4j.Variant.getMissing()
   * @param title Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(196) //= 0xc4. The runtime will prefer the VTID if present
  @VTID(58)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object gallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object plotBy,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seriesLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLegend,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object title);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter valueTitle is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter extraTitle is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param gallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param plotBy Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLegend Optional parameter. Default value is com4j.Variant.getMissing()
   * @param title Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryTitle Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(196) //= 0xc4. The runtime will prefer the VTID if present
  @VTID(58)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9, 10, 11}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object gallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object plotBy,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seriesLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLegend,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object title,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryTitle);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter extraTitle is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle, valueTitle, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param gallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param plotBy Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLegend Optional parameter. Default value is com4j.Variant.getMissing()
   * @param title Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryTitle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param valueTitle Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(196) //= 0xc4. The runtime will prefer the VTID if present
  @VTID(58)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, optParamIndex = {10, 11}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object gallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object plotBy,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seriesLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLegend,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object title,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryTitle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object valueTitle);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * chartWizard(source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle, valueTitle, extraTitle, 1033);
   * </code>
   * </p>
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param gallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param plotBy Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLegend Optional parameter. Default value is com4j.Variant.getMissing()
   * @param title Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryTitle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param valueTitle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param extraTitle Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(196) //= 0xc4. The runtime will prefer the VTID if present
  @VTID(58)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, optParamIndex = {11}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object gallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object plotBy,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seriesLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLegend,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object title,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryTitle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object valueTitle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object extraTitle);

  /**
   * @param source Optional parameter. Default value is com4j.Variant.getMissing()
   * @param gallery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param plotBy Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param seriesLabels Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLegend Optional parameter. Default value is com4j.Variant.getMissing()
   * @param title Optional parameter. Default value is com4j.Variant.getMissing()
   * @param categoryTitle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param valueTitle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param extraTitle Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(196) //= 0xc4. The runtime will prefer the VTID if present
  @VTID(58)
  void chartWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object gallery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object plotBy,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object seriesLabels,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLegend,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object title,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object categoryTitle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object valueTitle,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object extraTitle,
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
   * checkBoxes(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(824) //= 0x338. The runtime will prefer the VTID if present
  @VTID(59)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject checkBoxes();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkBoxes(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(824) //= 0x338. The runtime will prefer the VTID if present
  @VTID(59)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject checkBoxes(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(824) //= 0x338. The runtime will prefer the VTID if present
  @VTID(59)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject checkBoxes(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter customDictionary is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreUppercase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter alwaysSuggest is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter spellLang is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkSpelling(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(505) //= 0x1f9. The runtime will prefer the VTID if present
  @VTID(60)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void checkSpelling();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter ignoreUppercase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter alwaysSuggest is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter spellLang is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkSpelling(customDictionary, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(505) //= 0x1f9. The runtime will prefer the VTID if present
  @VTID(60)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void checkSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter alwaysSuggest is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter spellLang is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkSpelling(customDictionary, ignoreUppercase, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreUppercase Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(505) //= 0x1f9. The runtime will prefer the VTID if present
  @VTID(60)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void checkSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreUppercase);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter spellLang is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkSpelling(customDictionary, ignoreUppercase, alwaysSuggest, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreUppercase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alwaysSuggest Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(505) //= 0x1f9. The runtime will prefer the VTID if present
  @VTID(60)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void checkSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreUppercase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object alwaysSuggest);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkSpelling(customDictionary, ignoreUppercase, alwaysSuggest, spellLang, 1033);
   * </code>
   * </p>
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreUppercase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alwaysSuggest Optional parameter. Default value is com4j.Variant.getMissing()
   * @param spellLang Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(505) //= 0x1f9. The runtime will prefer the VTID if present
  @VTID(60)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void checkSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreUppercase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object alwaysSuggest,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object spellLang);

  /**
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreUppercase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alwaysSuggest Optional parameter. Default value is com4j.Variant.getMissing()
   * @param spellLang Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(505) //= 0x1f9. The runtime will prefer the VTID if present
  @VTID(60)
  void checkSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreUppercase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object alwaysSuggest,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object spellLang,
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
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartGroup
   */

  @DISPID(19) //= 0x13. The runtime will prefer the VTID if present
  @VTID(61)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.ChartGroup getColumn3DGroup();

  /**
   * <p>
   * Getter method for the COM property "Column3DGroup"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartGroup
   */

  @DISPID(19) //= 0x13. The runtime will prefer the VTID if present
  @VTID(61)
  com.exceljava.com4j.excel.ChartGroup getColumn3DGroup(
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
  @VTID(62)
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
  @VTID(62)
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
  @VTID(62)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject columnGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlPictureAppearance parameter appearance is set to 1</li><li>com.exceljava.com4j.excel.XlCopyPictureFormat parameter format is set to -4147</li><li>com.exceljava.com4j.excel.XlPictureAppearance parameter size is set to 2</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copyPicture(1, -4147, 2, 1033);
   * </code>
   * </p>
   */

  @DISPID(213) //= 0xd5. The runtime will prefer the VTID if present
  @VTID(63)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {com.exceljava.com4j.excel.XlPictureAppearance.class, com.exceljava.com4j.excel.XlCopyPictureFormat.class, com.exceljava.com4j.excel.XlPictureAppearance.class, int.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"1", "-4147", "2", "1033"})
  void copyPicture();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlCopyPictureFormat parameter format is set to -4147</li><li>com.exceljava.com4j.excel.XlPictureAppearance parameter size is set to 2</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copyPicture(appearance, -4147, 2, 1033);
   * </code>
   * </p>
   * @param appearance Optional parameter. Default value is 1
   */

  @DISPID(213) //= 0xd5. The runtime will prefer the VTID if present
  @VTID(63)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {com.exceljava.com4j.excel.XlCopyPictureFormat.class, com.exceljava.com4j.excel.XlPictureAppearance.class, int.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"-4147", "2", "1033"})
  void copyPicture(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlPictureAppearance parameter size is set to 2</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copyPicture(appearance, format, 2, 1033);
   * </code>
   * </p>
   * @param appearance Optional parameter. Default value is 1
   * @param format Optional parameter. Default value is -4147
   */

  @DISPID(213) //= 0xd5. The runtime will prefer the VTID if present
  @VTID(63)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {com.exceljava.com4j.excel.XlPictureAppearance.class, int.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"2", "1033"})
  void copyPicture(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @DefaultValue("-4147") com.exceljava.com4j.excel.XlCopyPictureFormat format);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
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

  @DISPID(213) //= 0xd5. The runtime will prefer the VTID if present
  @VTID(63)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void copyPicture(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @DefaultValue("-4147") com.exceljava.com4j.excel.XlCopyPictureFormat format,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlPictureAppearance size);

  /**
   * @param appearance Optional parameter. Default value is 1
   * @param format Optional parameter. Default value is -4147
   * @param size Optional parameter. Default value is 2
   * @param lcid Mandatory int parameter.
   */

  @DISPID(213) //= 0xd5. The runtime will prefer the VTID if present
  @VTID(63)
  void copyPicture(
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @DefaultValue("-4147") com.exceljava.com4j.excel.XlCopyPictureFormat format,
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlPictureAppearance size,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Corners"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getCorners(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Corners
   */

  @DISPID(79) //= 0x4f. The runtime will prefer the VTID if present
  @VTID(64)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.Corners getCorners();

  /**
   * <p>
   * Getter method for the COM property "Corners"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Corners
   */

  @DISPID(79) //= 0x4f. The runtime will prefer the VTID if present
  @VTID(64)
  com.exceljava.com4j.excel.Corners getCorners(
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter edition is set to com4j.Variant.getMissing()</li><li>com.exceljava.com4j.excel.XlPictureAppearance parameter appearance is set to 1</li><li>com.exceljava.com4j.excel.XlPictureAppearance parameter size is set to 1</li><li>java.lang.Object parameter containsPICT is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsBIFF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsRTF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsVALU is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createPublisher(com4j.Variant.getMissing(), 1, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(458) //= 0x1ca. The runtime will prefer the VTID if present
  @VTID(65)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, com.exceljava.com4j.excel.XlPictureAppearance.class, com.exceljava.com4j.excel.XlPictureAppearance.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1", "1", "80020004", "80020004", "80020004", "80020004", "1033"})
  void createPublisher();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlPictureAppearance parameter appearance is set to 1</li><li>com.exceljava.com4j.excel.XlPictureAppearance parameter size is set to 1</li><li>java.lang.Object parameter containsPICT is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsBIFF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsRTF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsVALU is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createPublisher(edition, 1, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param edition Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(458) //= 0x1ca. The runtime will prefer the VTID if present
  @VTID(65)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7}, javaType = {com.exceljava.com4j.excel.XlPictureAppearance.class, com.exceljava.com4j.excel.XlPictureAppearance.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.Int32, NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"1", "1", "80020004", "80020004", "80020004", "80020004", "1033"})
  void createPublisher(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object edition);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlPictureAppearance parameter size is set to 1</li><li>java.lang.Object parameter containsPICT is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsBIFF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsRTF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsVALU is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createPublisher(edition, appearance, 1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param edition Optional parameter. Default value is com4j.Variant.getMissing()
   * @param appearance Optional parameter. Default value is 1
   */

  @DISPID(458) //= 0x1ca. The runtime will prefer the VTID if present
  @VTID(65)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7}, javaType = {com.exceljava.com4j.excel.XlPictureAppearance.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"1", "80020004", "80020004", "80020004", "80020004", "1033"})
  void createPublisher(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object edition,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter containsPICT is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsBIFF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsRTF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsVALU is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createPublisher(edition, appearance, size, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param edition Optional parameter. Default value is com4j.Variant.getMissing()
   * @param appearance Optional parameter. Default value is 1
   * @param size Optional parameter. Default value is 1
   */

  @DISPID(458) //= 0x1ca. The runtime will prefer the VTID if present
  @VTID(65)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void createPublisher(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object edition,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance size);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter containsBIFF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsRTF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsVALU is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createPublisher(edition, appearance, size, containsPICT, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param edition Optional parameter. Default value is com4j.Variant.getMissing()
   * @param appearance Optional parameter. Default value is 1
   * @param size Optional parameter. Default value is 1
   * @param containsPICT Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(458) //= 0x1ca. The runtime will prefer the VTID if present
  @VTID(65)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void createPublisher(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object edition,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance size,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsPICT);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter containsRTF is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter containsVALU is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createPublisher(edition, appearance, size, containsPICT, containsBIFF, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param edition Optional parameter. Default value is com4j.Variant.getMissing()
   * @param appearance Optional parameter. Default value is 1
   * @param size Optional parameter. Default value is 1
   * @param containsPICT Optional parameter. Default value is com4j.Variant.getMissing()
   * @param containsBIFF Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(458) //= 0x1ca. The runtime will prefer the VTID if present
  @VTID(65)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void createPublisher(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object edition,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance size,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsPICT,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsBIFF);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter containsVALU is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createPublisher(edition, appearance, size, containsPICT, containsBIFF, containsRTF, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param edition Optional parameter. Default value is com4j.Variant.getMissing()
   * @param appearance Optional parameter. Default value is 1
   * @param size Optional parameter. Default value is 1
   * @param containsPICT Optional parameter. Default value is com4j.Variant.getMissing()
   * @param containsBIFF Optional parameter. Default value is com4j.Variant.getMissing()
   * @param containsRTF Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(458) //= 0x1ca. The runtime will prefer the VTID if present
  @VTID(65)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void createPublisher(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object edition,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance size,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsPICT,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsBIFF,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsRTF);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * createPublisher(edition, appearance, size, containsPICT, containsBIFF, containsRTF, containsVALU, 1033);
   * </code>
   * </p>
   * @param edition Optional parameter. Default value is com4j.Variant.getMissing()
   * @param appearance Optional parameter. Default value is 1
   * @param size Optional parameter. Default value is 1
   * @param containsPICT Optional parameter. Default value is com4j.Variant.getMissing()
   * @param containsBIFF Optional parameter. Default value is com4j.Variant.getMissing()
   * @param containsRTF Optional parameter. Default value is com4j.Variant.getMissing()
   * @param containsVALU Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(458) //= 0x1ca. The runtime will prefer the VTID if present
  @VTID(65)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void createPublisher(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object edition,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance size,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsPICT,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsBIFF,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsRTF,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsVALU);

  /**
   * @param edition Optional parameter. Default value is com4j.Variant.getMissing()
   * @param appearance Optional parameter. Default value is 1
   * @param size Optional parameter. Default value is 1
   * @param containsPICT Optional parameter. Default value is com4j.Variant.getMissing()
   * @param containsBIFF Optional parameter. Default value is com4j.Variant.getMissing()
   * @param containsRTF Optional parameter. Default value is com4j.Variant.getMissing()
   * @param containsVALU Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(458) //= 0x1ca. The runtime will prefer the VTID if present
  @VTID(65)
  void createPublisher(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object edition,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @DefaultValue("1") com.exceljava.com4j.excel.XlPictureAppearance size,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsPICT,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsBIFF,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsRTF,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object containsVALU,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "DataTable"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.DataTable
   */

  @DISPID(1395) //= 0x573. The runtime will prefer the VTID if present
  @VTID(66)
  com.exceljava.com4j.excel.DataTable getDataTable();


  /**
   * <p>
   * Getter method for the COM property "DepthPercent"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getDepthPercent(1033);
   * </code>
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(48) //= 0x30. The runtime will prefer the VTID if present
  @VTID(67)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  int getDepthPercent();

  /**
   * <p>
   * Getter method for the COM property "DepthPercent"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type int
   */

  @DISPID(48) //= 0x30. The runtime will prefer the VTID if present
  @VTID(67)
  int getDepthPercent(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "DepthPercent"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setDepthPercent(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(48) //= 0x30. The runtime will prefer the VTID if present
  @VTID(68)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setDepthPercent(
    int rhs);

  /**
   * <p>
   * Setter method for the COM property "DepthPercent"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory int parameter.
   */

  @DISPID(48) //= 0x30. The runtime will prefer the VTID if present
  @VTID(68)
  void setDepthPercent(
    @LCID int lcid,
    int rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * deselect(1033);
   * </code>
   * </p>
   */

  @DISPID(1120) //= 0x460. The runtime will prefer the VTID if present
  @VTID(69)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void deselect();

  /**
   * @param lcid Mandatory int parameter.
   */

  @DISPID(1120) //= 0x460. The runtime will prefer the VTID if present
  @VTID(69)
  void deselect(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "DisplayBlanksAs"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getDisplayBlanksAs(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlDisplayBlanksAs
   */

  @DISPID(93) //= 0x5d. The runtime will prefer the VTID if present
  @VTID(70)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.XlDisplayBlanksAs getDisplayBlanksAs();

  /**
   * <p>
   * Getter method for the COM property "DisplayBlanksAs"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.XlDisplayBlanksAs
   */

  @DISPID(93) //= 0x5d. The runtime will prefer the VTID if present
  @VTID(70)
  com.exceljava.com4j.excel.XlDisplayBlanksAs getDisplayBlanksAs(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "DisplayBlanksAs"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setDisplayBlanksAs(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlDisplayBlanksAs parameter.
   */

  @DISPID(93) //= 0x5d. The runtime will prefer the VTID if present
  @VTID(71)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setDisplayBlanksAs(
    com.exceljava.com4j.excel.XlDisplayBlanksAs rhs);

  /**
   * <p>
   * Setter method for the COM property "DisplayBlanksAs"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory com.exceljava.com4j.excel.XlDisplayBlanksAs parameter.
   */

  @DISPID(93) //= 0x5d. The runtime will prefer the VTID if present
  @VTID(71)
  void setDisplayBlanksAs(
    @LCID int lcid,
    com.exceljava.com4j.excel.XlDisplayBlanksAs rhs);


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
  @VTID(72)
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
  @VTID(72)
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
  @VTID(72)
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
   * drawings(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(772) //= 0x304. The runtime will prefer the VTID if present
  @VTID(73)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject drawings();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * drawings(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(772) //= 0x304. The runtime will prefer the VTID if present
  @VTID(73)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject drawings(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(772) //= 0x304. The runtime will prefer the VTID if present
  @VTID(73)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject drawings(
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
   * drawingObjects(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(88) //= 0x58. The runtime will prefer the VTID if present
  @VTID(74)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject drawingObjects();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * drawingObjects(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(88) //= 0x58. The runtime will prefer the VTID if present
  @VTID(74)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject drawingObjects(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(88) //= 0x58. The runtime will prefer the VTID if present
  @VTID(74)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject drawingObjects(
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
   * dropDowns(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(836) //= 0x344. The runtime will prefer the VTID if present
  @VTID(75)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject dropDowns();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * dropDowns(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(836) //= 0x344. The runtime will prefer the VTID if present
  @VTID(75)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject dropDowns(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(836) //= 0x344. The runtime will prefer the VTID if present
  @VTID(75)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject dropDowns(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Elevation"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getElevation(1033);
   * </code>
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(49) //= 0x31. The runtime will prefer the VTID if present
  @VTID(76)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  int getElevation();

  /**
   * <p>
   * Getter method for the COM property "Elevation"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type int
   */

  @DISPID(49) //= 0x31. The runtime will prefer the VTID if present
  @VTID(76)
  int getElevation(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Elevation"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setElevation(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(49) //= 0x31. The runtime will prefer the VTID if present
  @VTID(77)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setElevation(
    int rhs);

  /**
   * <p>
   * Setter method for the COM property "Elevation"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory int parameter.
   */

  @DISPID(49) //= 0x31. The runtime will prefer the VTID if present
  @VTID(77)
  void setElevation(
    @LCID int lcid,
    int rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * evaluate(name, 1033);
   * </code>
   * </p>
   * @param name Mandatory java.lang.Object parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(78)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object evaluate(
    @MarshalAs(NativeType.VARIANT) java.lang.Object name);

  /**
   * @param name Mandatory java.lang.Object parameter.
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(78)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object evaluate(
    @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _Evaluate(name, 1033);
   * </code>
   * </p>
   * @param name Mandatory java.lang.Object parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(-5) //= 0xfffffffb. The runtime will prefer the VTID if present
  @VTID(79)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object _Evaluate(
    @MarshalAs(NativeType.VARIANT) java.lang.Object name);

  /**
   * @param name Mandatory java.lang.Object parameter.
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(-5) //= 0xfffffffb. The runtime will prefer the VTID if present
  @VTID(79)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object _Evaluate(
    @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Floor"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getFloor(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Floor
   */

  @DISPID(83) //= 0x53. The runtime will prefer the VTID if present
  @VTID(80)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.Floor getFloor();

  /**
   * <p>
   * Getter method for the COM property "Floor"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Floor
   */

  @DISPID(83) //= 0x53. The runtime will prefer the VTID if present
  @VTID(80)
  com.exceljava.com4j.excel.Floor getFloor(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "GapDepth"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getGapDepth(1033);
   * </code>
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(50) //= 0x32. The runtime will prefer the VTID if present
  @VTID(81)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  int getGapDepth();

  /**
   * <p>
   * Getter method for the COM property "GapDepth"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type int
   */

  @DISPID(50) //= 0x32. The runtime will prefer the VTID if present
  @VTID(81)
  int getGapDepth(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "GapDepth"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setGapDepth(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(50) //= 0x32. The runtime will prefer the VTID if present
  @VTID(82)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setGapDepth(
    int rhs);

  /**
   * <p>
   * Setter method for the COM property "GapDepth"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory int parameter.
   */

  @DISPID(50) //= 0x32. The runtime will prefer the VTID if present
  @VTID(82)
  void setGapDepth(
    @LCID int lcid,
    int rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * groupBoxes(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(834) //= 0x342. The runtime will prefer the VTID if present
  @VTID(83)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject groupBoxes();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * groupBoxes(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(834) //= 0x342. The runtime will prefer the VTID if present
  @VTID(83)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject groupBoxes(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(834) //= 0x342. The runtime will prefer the VTID if present
  @VTID(83)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject groupBoxes(
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
   * groupObjects(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1113) //= 0x459. The runtime will prefer the VTID if present
  @VTID(84)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject groupObjects();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * groupObjects(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1113) //= 0x459. The runtime will prefer the VTID if present
  @VTID(84)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject groupObjects(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1113) //= 0x459. The runtime will prefer the VTID if present
  @VTID(84)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject groupObjects(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "HasAxis"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter index2 is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getHasAxis(com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(52) //= 0x34. The runtime will prefer the VTID if present
  @VTID(85)
  @UseDefaultValues(paramIndexMapping = {3}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  @ReturnValue(type=NativeType.VARIANT,index=3)
  java.lang.Object getHasAxis();

  /**
   * <p>
   * Getter method for the COM property "HasAxis"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index2 is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getHasAxis(index1, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param index1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(52) //= 0x34. The runtime will prefer the VTID if present
  @VTID(85)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.VARIANT,index=3)
  java.lang.Object getHasAxis(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index1);

  /**
   * <p>
   * Getter method for the COM property "HasAxis"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getHasAxis(index1, index2, 1033);
   * </code>
   * </p>
   * @param index1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param index2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(52) //= 0x34. The runtime will prefer the VTID if present
  @VTID(85)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=3)
  java.lang.Object getHasAxis(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index2);

  /**
   * <p>
   * Getter method for the COM property "HasAxis"
   * </p>
   * @param index1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param index2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(52) //= 0x34. The runtime will prefer the VTID if present
  @VTID(85)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getHasAxis(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index2,
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "HasAxis"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index1 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter index2 is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setHasAxis(com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(52) //= 0x34. The runtime will prefer the VTID if present
  @VTID(86)
  @UseDefaultValues(paramIndexMapping = {3}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void setHasAxis(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "HasAxis"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index2 is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setHasAxis(index1, com4j.Variant.getMissing(), 1033, rhs);
   * </code>
   * </p>
   * @param index1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(52) //= 0x34. The runtime will prefer the VTID if present
  @VTID(86)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void setHasAxis(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index1,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "HasAxis"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setHasAxis(index1, index2, 1033, rhs);
   * </code>
   * </p>
   * @param index1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param index2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(52) //= 0x34. The runtime will prefer the VTID if present
  @VTID(86)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setHasAxis(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index2,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "HasAxis"
   * </p>
   * @param index1 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param index2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(52) //= 0x34. The runtime will prefer the VTID if present
  @VTID(86)
  void setHasAxis(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index2,
    @LCID int lcid,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "HasDataTable"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1396) //= 0x574. The runtime will prefer the VTID if present
  @VTID(87)
  boolean getHasDataTable();


  /**
   * <p>
   * Setter method for the COM property "HasDataTable"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1396) //= 0x574. The runtime will prefer the VTID if present
  @VTID(88)
  void setHasDataTable(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasLegend"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getHasLegend(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(53) //= 0x35. The runtime will prefer the VTID if present
  @VTID(89)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getHasLegend();

  /**
   * <p>
   * Getter method for the COM property "HasLegend"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(53) //= 0x35. The runtime will prefer the VTID if present
  @VTID(89)
  boolean getHasLegend(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "HasLegend"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setHasLegend(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(53) //= 0x35. The runtime will prefer the VTID if present
  @VTID(90)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setHasLegend(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "HasLegend"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(53) //= 0x35. The runtime will prefer the VTID if present
  @VTID(90)
  void setHasLegend(
    @LCID int lcid,
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HasTitle"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getHasTitle(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(54) //= 0x36. The runtime will prefer the VTID if present
  @VTID(91)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getHasTitle();

  /**
   * <p>
   * Getter method for the COM property "HasTitle"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(54) //= 0x36. The runtime will prefer the VTID if present
  @VTID(91)
  boolean getHasTitle(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "HasTitle"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setHasTitle(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(54) //= 0x36. The runtime will prefer the VTID if present
  @VTID(92)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setHasTitle(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "HasTitle"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(54) //= 0x36. The runtime will prefer the VTID if present
  @VTID(92)
  void setHasTitle(
    @LCID int lcid,
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "HeightPercent"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getHeightPercent(1033);
   * </code>
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(55) //= 0x37. The runtime will prefer the VTID if present
  @VTID(93)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  int getHeightPercent();

  /**
   * <p>
   * Getter method for the COM property "HeightPercent"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type int
   */

  @DISPID(55) //= 0x37. The runtime will prefer the VTID if present
  @VTID(93)
  int getHeightPercent(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "HeightPercent"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setHeightPercent(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(55) //= 0x37. The runtime will prefer the VTID if present
  @VTID(94)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setHeightPercent(
    int rhs);

  /**
   * <p>
   * Setter method for the COM property "HeightPercent"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory int parameter.
   */

  @DISPID(55) //= 0x37. The runtime will prefer the VTID if present
  @VTID(94)
  void setHeightPercent(
    @LCID int lcid,
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Hyperlinks"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Hyperlinks
   */

  @DISPID(1393) //= 0x571. The runtime will prefer the VTID if present
  @VTID(95)
  com.exceljava.com4j.excel.Hyperlinks getHyperlinks();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * labels(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(841) //= 0x349. The runtime will prefer the VTID if present
  @VTID(96)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject labels();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * labels(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(841) //= 0x349. The runtime will prefer the VTID if present
  @VTID(96)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject labels(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(841) //= 0x349. The runtime will prefer the VTID if present
  @VTID(96)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject labels(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Legend"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getLegend(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Legend
   */

  @DISPID(84) //= 0x54. The runtime will prefer the VTID if present
  @VTID(97)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.Legend getLegend();

  /**
   * <p>
   * Getter method for the COM property "Legend"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Legend
   */

  @DISPID(84) //= 0x54. The runtime will prefer the VTID if present
  @VTID(97)
  com.exceljava.com4j.excel.Legend getLegend(
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
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartGroup
   */

  @DISPID(20) //= 0x14. The runtime will prefer the VTID if present
  @VTID(98)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.ChartGroup getLine3DGroup();

  /**
   * <p>
   * Getter method for the COM property "Line3DGroup"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartGroup
   */

  @DISPID(20) //= 0x14. The runtime will prefer the VTID if present
  @VTID(98)
  com.exceljava.com4j.excel.ChartGroup getLine3DGroup(
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
  @VTID(99)
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
  @VTID(99)
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
  @VTID(99)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject lineGroups(
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
   * lines(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(767) //= 0x2ff. The runtime will prefer the VTID if present
  @VTID(100)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject lines();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * lines(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(767) //= 0x2ff. The runtime will prefer the VTID if present
  @VTID(100)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject lines(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(767) //= 0x2ff. The runtime will prefer the VTID if present
  @VTID(100)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject lines(
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
   * listBoxes(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(832) //= 0x340. The runtime will prefer the VTID if present
  @VTID(101)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject listBoxes();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * listBoxes(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(832) //= 0x340. The runtime will prefer the VTID if present
  @VTID(101)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject listBoxes(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(832) //= 0x340. The runtime will prefer the VTID if present
  @VTID(101)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject listBoxes(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter name is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * location(where, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param where Mandatory com.exceljava.com4j.excel.XlChartLocation parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel._Chart
   */

  @DISPID(1397) //= 0x575. The runtime will prefer the VTID if present
  @VTID(102)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel._Chart location(
    com.exceljava.com4j.excel.XlChartLocation where);

  /**
   * @param where Mandatory com.exceljava.com4j.excel.XlChartLocation parameter.
   * @param name Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel._Chart
   */

  @DISPID(1397) //= 0x575. The runtime will prefer the VTID if present
  @VTID(102)
  com.exceljava.com4j.excel._Chart location(
    com.exceljava.com4j.excel.XlChartLocation where,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object name);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * oleObjects(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(799) //= 0x31f. The runtime will prefer the VTID if present
  @VTID(103)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject oleObjects();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * oleObjects(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(799) //= 0x31f. The runtime will prefer the VTID if present
  @VTID(103)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject oleObjects(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(799) //= 0x31f. The runtime will prefer the VTID if present
  @VTID(103)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject oleObjects(
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
   * optionButtons(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(826) //= 0x33a. The runtime will prefer the VTID if present
  @VTID(104)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject optionButtons();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * optionButtons(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(826) //= 0x33a. The runtime will prefer the VTID if present
  @VTID(104)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject optionButtons(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(826) //= 0x33a. The runtime will prefer the VTID if present
  @VTID(104)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject optionButtons(
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
   * ovals(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(801) //= 0x321. The runtime will prefer the VTID if present
  @VTID(105)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject ovals();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * ovals(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(801) //= 0x321. The runtime will prefer the VTID if present
  @VTID(105)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject ovals(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(801) //= 0x321. The runtime will prefer the VTID if present
  @VTID(105)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject ovals(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * paste(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(211) //= 0xd3. The runtime will prefer the VTID if present
  @VTID(106)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void paste();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * paste(type, 1033);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(211) //= 0xd3. The runtime will prefer the VTID if present
  @VTID(106)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void paste(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type);

  /**
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(211) //= 0xd3. The runtime will prefer the VTID if present
  @VTID(106)
  void paste(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Perspective"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getPerspective(1033);
   * </code>
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(57) //= 0x39. The runtime will prefer the VTID if present
  @VTID(107)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  int getPerspective();

  /**
   * <p>
   * Getter method for the COM property "Perspective"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type int
   */

  @DISPID(57) //= 0x39. The runtime will prefer the VTID if present
  @VTID(107)
  int getPerspective(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Perspective"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setPerspective(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(57) //= 0x39. The runtime will prefer the VTID if present
  @VTID(108)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setPerspective(
    int rhs);

  /**
   * <p>
   * Setter method for the COM property "Perspective"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory int parameter.
   */

  @DISPID(57) //= 0x39. The runtime will prefer the VTID if present
  @VTID(108)
  void setPerspective(
    @LCID int lcid,
    int rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pictures(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(771) //= 0x303. The runtime will prefer the VTID if present
  @VTID(109)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject pictures();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pictures(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(771) //= 0x303. The runtime will prefer the VTID if present
  @VTID(109)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject pictures(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(771) //= 0x303. The runtime will prefer the VTID if present
  @VTID(109)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject pictures(
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
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartGroup
   */

  @DISPID(21) //= 0x15. The runtime will prefer the VTID if present
  @VTID(110)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.ChartGroup getPie3DGroup();

  /**
   * <p>
   * Getter method for the COM property "Pie3DGroup"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartGroup
   */

  @DISPID(21) //= 0x15. The runtime will prefer the VTID if present
  @VTID(110)
  com.exceljava.com4j.excel.ChartGroup getPie3DGroup(
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
  @VTID(111)
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
  @VTID(111)
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
  @VTID(111)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject pieGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "PlotArea"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getPlotArea(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PlotArea
   */

  @DISPID(85) //= 0x55. The runtime will prefer the VTID if present
  @VTID(112)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.PlotArea getPlotArea();

  /**
   * <p>
   * Getter method for the COM property "PlotArea"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.PlotArea
   */

  @DISPID(85) //= 0x55. The runtime will prefer the VTID if present
  @VTID(112)
  com.exceljava.com4j.excel.PlotArea getPlotArea(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "PlotVisibleOnly"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getPlotVisibleOnly(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(92) //= 0x5c. The runtime will prefer the VTID if present
  @VTID(113)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getPlotVisibleOnly();

  /**
   * <p>
   * Getter method for the COM property "PlotVisibleOnly"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(92) //= 0x5c. The runtime will prefer the VTID if present
  @VTID(113)
  boolean getPlotVisibleOnly(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "PlotVisibleOnly"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setPlotVisibleOnly(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(92) //= 0x5c. The runtime will prefer the VTID if present
  @VTID(114)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setPlotVisibleOnly(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "PlotVisibleOnly"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(92) //= 0x5c. The runtime will prefer the VTID if present
  @VTID(114)
  void setPlotVisibleOnly(
    @LCID int lcid,
    boolean rhs);


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
  @VTID(115)
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
  @VTID(115)
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
  @VTID(115)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject radarGroups(
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
   * rectangles(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(774) //= 0x306. The runtime will prefer the VTID if present
  @VTID(116)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject rectangles();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * rectangles(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(774) //= 0x306. The runtime will prefer the VTID if present
  @VTID(116)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject rectangles(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(774) //= 0x306. The runtime will prefer the VTID if present
  @VTID(116)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject rectangles(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "RightAngleAxes"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getRightAngleAxes(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(58) //= 0x3a. The runtime will prefer the VTID if present
  @VTID(117)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getRightAngleAxes();

  /**
   * <p>
   * Getter method for the COM property "RightAngleAxes"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(58) //= 0x3a. The runtime will prefer the VTID if present
  @VTID(117)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getRightAngleAxes(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "RightAngleAxes"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setRightAngleAxes(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(58) //= 0x3a. The runtime will prefer the VTID if present
  @VTID(118)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setRightAngleAxes(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "RightAngleAxes"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(58) //= 0x3a. The runtime will prefer the VTID if present
  @VTID(118)
  void setRightAngleAxes(
    @LCID int lcid,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Rotation"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getRotation(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(59) //= 0x3b. The runtime will prefer the VTID if present
  @VTID(119)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getRotation();

  /**
   * <p>
   * Getter method for the COM property "Rotation"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(59) //= 0x3b. The runtime will prefer the VTID if present
  @VTID(119)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getRotation(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Rotation"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setRotation(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(59) //= 0x3b. The runtime will prefer the VTID if present
  @VTID(120)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setRotation(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);

  /**
   * <p>
   * Setter method for the COM property "Rotation"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(59) //= 0x3b. The runtime will prefer the VTID if present
  @VTID(120)
  void setRotation(
    @LCID int lcid,
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * scrollBars(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(830) //= 0x33e. The runtime will prefer the VTID if present
  @VTID(121)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject scrollBars();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * scrollBars(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(830) //= 0x33e. The runtime will prefer the VTID if present
  @VTID(121)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject scrollBars(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(830) //= 0x33e. The runtime will prefer the VTID if present
  @VTID(121)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject scrollBars(
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
   * seriesCollection(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(68) //= 0x44. The runtime will prefer the VTID if present
  @VTID(122)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject seriesCollection();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * seriesCollection(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(68) //= 0x44. The runtime will prefer the VTID if present
  @VTID(122)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject seriesCollection(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(68) //= 0x44. The runtime will prefer the VTID if present
  @VTID(122)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject seriesCollection(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "SizeWithWindow"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getSizeWithWindow(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(94) //= 0x5e. The runtime will prefer the VTID if present
  @VTID(123)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getSizeWithWindow();

  /**
   * <p>
   * Getter method for the COM property "SizeWithWindow"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(94) //= 0x5e. The runtime will prefer the VTID if present
  @VTID(123)
  boolean getSizeWithWindow(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "SizeWithWindow"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setSizeWithWindow(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(94) //= 0x5e. The runtime will prefer the VTID if present
  @VTID(124)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setSizeWithWindow(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "SizeWithWindow"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(94) //= 0x5e. The runtime will prefer the VTID if present
  @VTID(124)
  void setSizeWithWindow(
    @LCID int lcid,
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowWindow"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1399) //= 0x577. The runtime will prefer the VTID if present
  @VTID(125)
  boolean getShowWindow();


  /**
   * <p>
   * Setter method for the COM property "ShowWindow"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1399) //= 0x577. The runtime will prefer the VTID if present
  @VTID(126)
  void setShowWindow(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * spinners(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(838) //= 0x346. The runtime will prefer the VTID if present
  @VTID(127)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject spinners();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * spinners(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(838) //= 0x346. The runtime will prefer the VTID if present
  @VTID(127)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject spinners(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(838) //= 0x346. The runtime will prefer the VTID if present
  @VTID(127)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject spinners(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "SubType"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getSubType(1033);
   * </code>
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(128)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  int getSubType();

  /**
   * <p>
   * Getter method for the COM property "SubType"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type int
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(128)
  int getSubType(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "SubType"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setSubType(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(129)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setSubType(
    int rhs);

  /**
   * <p>
   * Setter method for the COM property "SubType"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory int parameter.
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(129)
  void setSubType(
    @LCID int lcid,
    int rhs);


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
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartGroup
   */

  @DISPID(22) //= 0x16. The runtime will prefer the VTID if present
  @VTID(130)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.ChartGroup getSurfaceGroup();

  /**
   * <p>
   * Getter method for the COM property "SurfaceGroup"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartGroup
   */

  @DISPID(22) //= 0x16. The runtime will prefer the VTID if present
  @VTID(130)
  com.exceljava.com4j.excel.ChartGroup getSurfaceGroup(
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
   * textBoxes(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(777) //= 0x309. The runtime will prefer the VTID if present
  @VTID(131)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject textBoxes();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * textBoxes(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(777) //= 0x309. The runtime will prefer the VTID if present
  @VTID(131)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject textBoxes(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(777) //= 0x309. The runtime will prefer the VTID if present
  @VTID(131)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject textBoxes(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getType(1033);
   * </code>
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(132)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  int getType();

  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type int
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(132)
  int getType(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "Type"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setType(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(133)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setType(
    int rhs);

  /**
   * <p>
   * Setter method for the COM property "Type"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory int parameter.
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(133)
  void setType(
    @LCID int lcid,
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "ChartType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlChartType
   */

  @DISPID(1400) //= 0x578. The runtime will prefer the VTID if present
  @VTID(134)
  com.exceljava.com4j.excel.XlChartType getChartType();


  /**
   * <p>
   * Setter method for the COM property "ChartType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlChartType parameter.
   */

  @DISPID(1400) //= 0x578. The runtime will prefer the VTID if present
  @VTID(135)
  void setChartType(
    com.exceljava.com4j.excel.XlChartType rhs);


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
   * @param chartType Mandatory com.exceljava.com4j.excel.XlChartType parameter.
   */

  @DISPID(1401) //= 0x579. The runtime will prefer the VTID if present
  @VTID(136)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void applyCustomType(
    com.exceljava.com4j.excel.XlChartType chartType);

  /**
   * @param chartType Mandatory com.exceljava.com4j.excel.XlChartType parameter.
   * @param typeName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1401) //= 0x579. The runtime will prefer the VTID if present
  @VTID(136)
  void applyCustomType(
    com.exceljava.com4j.excel.XlChartType chartType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object typeName);


  /**
   * <p>
   * Getter method for the COM property "Walls"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getWalls(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Walls
   */

  @DISPID(86) //= 0x56. The runtime will prefer the VTID if present
  @VTID(137)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.Walls getWalls();

  /**
   * <p>
   * Getter method for the COM property "Walls"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Walls
   */

  @DISPID(86) //= 0x56. The runtime will prefer the VTID if present
  @VTID(137)
  com.exceljava.com4j.excel.Walls getWalls(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "WallsAndGridlines2D"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getWallsAndGridlines2D(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(210) //= 0xd2. The runtime will prefer the VTID if present
  @VTID(138)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getWallsAndGridlines2D();

  /**
   * <p>
   * Getter method for the COM property "WallsAndGridlines2D"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(210) //= 0xd2. The runtime will prefer the VTID if present
  @VTID(138)
  boolean getWallsAndGridlines2D(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "WallsAndGridlines2D"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setWallsAndGridlines2D(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(210) //= 0xd2. The runtime will prefer the VTID if present
  @VTID(139)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setWallsAndGridlines2D(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "WallsAndGridlines2D"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(210) //= 0xd2. The runtime will prefer the VTID if present
  @VTID(139)
  void setWallsAndGridlines2D(
    @LCID int lcid,
    boolean rhs);


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
  @VTID(140)
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
  @VTID(140)
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
  @VTID(140)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject xyGroups(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "BarShape"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlBarShape
   */

  @DISPID(1403) //= 0x57b. The runtime will prefer the VTID if present
  @VTID(141)
  com.exceljava.com4j.excel.XlBarShape getBarShape();


  /**
   * <p>
   * Setter method for the COM property "BarShape"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlBarShape parameter.
   */

  @DISPID(1403) //= 0x57b. The runtime will prefer the VTID if present
  @VTID(142)
  void setBarShape(
    com.exceljava.com4j.excel.XlBarShape rhs);


  /**
   * <p>
   * Getter method for the COM property "PlotBy"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlRowCol
   */

  @DISPID(202) //= 0xca. The runtime will prefer the VTID if present
  @VTID(143)
  com.exceljava.com4j.excel.XlRowCol getPlotBy();


  /**
   * <p>
   * Setter method for the COM property "PlotBy"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlRowCol parameter.
   */

  @DISPID(202) //= 0xca. The runtime will prefer the VTID if present
  @VTID(144)
  void setPlotBy(
    com.exceljava.com4j.excel.XlRowCol rhs);


  /**
   */

  @DISPID(1404) //= 0x57c. The runtime will prefer the VTID if present
  @VTID(145)
  void copyChartBuild();


  /**
   * <p>
   * Getter method for the COM property "ProtectFormatting"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1405) //= 0x57d. The runtime will prefer the VTID if present
  @VTID(146)
  boolean getProtectFormatting();


  /**
   * <p>
   * Setter method for the COM property "ProtectFormatting"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1405) //= 0x57d. The runtime will prefer the VTID if present
  @VTID(147)
  void setProtectFormatting(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ProtectData"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1406) //= 0x57e. The runtime will prefer the VTID if present
  @VTID(148)
  boolean getProtectData();


  /**
   * <p>
   * Setter method for the COM property "ProtectData"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1406) //= 0x57e. The runtime will prefer the VTID if present
  @VTID(149)
  void setProtectData(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ProtectGoalSeek"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1407) //= 0x57f. The runtime will prefer the VTID if present
  @VTID(150)
  boolean getProtectGoalSeek();


  /**
   * <p>
   * Setter method for the COM property "ProtectGoalSeek"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1407) //= 0x57f. The runtime will prefer the VTID if present
  @VTID(151)
  void setProtectGoalSeek(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ProtectSelection"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1408) //= 0x580. The runtime will prefer the VTID if present
  @VTID(152)
  boolean getProtectSelection();


  /**
   * <p>
   * Setter method for the COM property "ProtectSelection"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1408) //= 0x580. The runtime will prefer the VTID if present
  @VTID(153)
  void setProtectSelection(
    boolean rhs);


  /**
   * @param x Mandatory int parameter.
   * @param y Mandatory int parameter.
   * @param elementID Mandatory Holder<Integer> parameter.
   * @param arg1 Mandatory Holder<Integer> parameter.
   * @param arg2 Mandatory Holder<Integer> parameter.
   */

  @DISPID(1409) //= 0x581. The runtime will prefer the VTID if present
  @VTID(154)
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
   * @param source Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(1413) //= 0x585. The runtime will prefer the VTID if present
  @VTID(155)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setSourceData(
    com.exceljava.com4j.excel.Range source);

  /**
   * @param source Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param plotBy Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1413) //= 0x585. The runtime will prefer the VTID if present
  @VTID(155)
  void setSourceData(
    com.exceljava.com4j.excel.Range source,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object plotBy);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter filterName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter interactive is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * export(filename, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(1414) //= 0x586. The runtime will prefer the VTID if present
  @VTID(156)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=3)
  boolean export(
    java.lang.String filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter interactive is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * export(filename, filterName, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param filterName Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @DISPID(1414) //= 0x586. The runtime will prefer the VTID if present
  @VTID(156)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  boolean export(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filterName);

  /**
   * @param filename Mandatory java.lang.String parameter.
   * @param filterName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param interactive Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type boolean
   */

  @DISPID(1414) //= 0x586. The runtime will prefer the VTID if present
  @VTID(156)
  boolean export(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filterName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object interactive);


  /**
   */

  @DISPID(1417) //= 0x589. The runtime will prefer the VTID if present
  @VTID(157)
  void refresh();


  /**
   * <p>
   * Getter method for the COM property "PivotLayout"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotLayout
   */

  @DISPID(1814) //= 0x716. The runtime will prefer the VTID if present
  @VTID(158)
  com.exceljava.com4j.excel.PivotLayout getPivotLayout();


  /**
   * <p>
   * Getter method for the COM property "HasPivotFields"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1815) //= 0x717. The runtime will prefer the VTID if present
  @VTID(159)
  boolean getHasPivotFields();


  /**
   * <p>
   * Setter method for the COM property "HasPivotFields"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1815) //= 0x717. The runtime will prefer the VTID if present
  @VTID(160)
  void setHasPivotFields(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Scripts"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Scripts
   */

  @DISPID(1816) //= 0x718. The runtime will prefer the VTID if present
  @VTID(161)
  com.exceljava.com4j.office.Scripts getScripts();


  @VTID(161)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.Scripts.class})
  com.exceljava.com4j.office.Script getScripts(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _PrintOut();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, preview, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, preview, activePrinter, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, preview, activePrinter, printToFile, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, preview, activePrinter, printToFile, collate, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PrintOut(from, to, copies, preview, activePrinter, printToFile, collate, prToFileName, 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @param prToFileName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(162)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object prToFileName);

  /**
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @param prToFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(1772) //= 0x6ec. The runtime will prefer the VTID if present
  @VTID(162)
  void _PrintOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object prToFileName,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Tab"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Tab
   */

  @DISPID(1041) //= 0x411. The runtime will prefer the VTID if present
  @VTID(163)
  com.exceljava.com4j.excel.Tab getTab();


  /**
   * <p>
   * Getter method for the COM property "MailEnvelope"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoEnvelopeVB
   */

  @DISPID(2021) //= 0x7e5. The runtime will prefer the VTID if present
  @VTID(164)
  com.exceljava.com4j.office.IMsoEnvelopeVB getMailEnvelope();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlDataLabelsType parameter type is set to 2</li><li>java.lang.Object parameter legendKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showSeriesName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showCategoryName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(165)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, javaType = {com.exceljava.com4j.excel.XlDataLabelsType.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.Int32, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"2", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void applyDataLabels();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter legendKey is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showSeriesName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showCategoryName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(165)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter autoText is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showSeriesName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showCategoryName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, legendKey, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(165)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter hasLeaderLines is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showSeriesName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showCategoryName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, legendKey, autoText, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(165)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter showSeriesName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showCategoryName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, legendKey, autoText, hasLeaderLines, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(165)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter showCategoryName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, legendKey, autoText, hasLeaderLines, showSeriesName, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showSeriesName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(165)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showSeriesName);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter showValue is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showSeriesName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showCategoryName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(165)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showSeriesName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showCategoryName);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter showPercentage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showSeriesName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showCategoryName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showValue Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(165)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
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
     * <li>java.lang.Object parameter showBubbleSize is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showSeriesName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showCategoryName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showValue Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showPercentage Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(165)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9, 10}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
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
     * <li>java.lang.Object parameter separator is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showSeriesName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showCategoryName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showValue Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showPercentage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showBubbleSize Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(165)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9, 10}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showSeriesName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showCategoryName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showValue,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showPercentage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showBubbleSize);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyDataLabels(type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize, separator, 1033);
   * </code>
   * </p>
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
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
  @VTID(165)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, optParamIndex = {10}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showSeriesName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showCategoryName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showValue,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showPercentage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showBubbleSize,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object separator);

  /**
   * @param type Optional parameter. Default value is 2
   * @param legendKey Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoText Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasLeaderLines Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showSeriesName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showCategoryName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showValue Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showPercentage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param showBubbleSize Optional parameter. Default value is com4j.Variant.getMissing()
   * @param separator Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(1922) //= 0x782. The runtime will prefer the VTID if present
  @VTID(165)
  void applyDataLabels(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlDataLabelsType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object legendKey,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoText,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasLeaderLines,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showSeriesName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showCategoryName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showValue,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showPercentage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object showBubbleSize,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object separator,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter fileFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(166)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void _SaveAs(
    java.lang.String filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(166)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void _SaveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, password, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(166)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void _SaveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, password, writeResPassword, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(166)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void _SaveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(166)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  void _SaveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(166)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void _SaveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(166)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void _SaveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(166)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void _SaveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _SaveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage, textVisualLayout, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(166)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void _SaveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout);

  /**
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @param local Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1925) //= 0x785. The runtime will prefer the VTID if present
  @VTID(166)
  void _SaveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object local);


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

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(167)
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

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(167)
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

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(167)
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

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(167)
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

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(167)
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

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(167)
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scenarios,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object userInterfaceOnly);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter chartType is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * applyLayout(layout, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param layout Mandatory int parameter.
   */

  @DISPID(2500) //= 0x9c4. The runtime will prefer the VTID if present
  @VTID(168)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void applyLayout(
    int layout);

  /**
   * @param layout Mandatory int parameter.
   * @param chartType Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2500) //= 0x9c4. The runtime will prefer the VTID if present
  @VTID(168)
  void applyLayout(
    int layout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object chartType);


  /**
   * @param element Mandatory com.exceljava.com4j.office.MsoChartElementType parameter.
   */

  @DISPID(2502) //= 0x9c6. The runtime will prefer the VTID if present
  @VTID(169)
  void setElement(
    com.exceljava.com4j.office.MsoChartElementType element);


  /**
   * <p>
   * Getter method for the COM property "ShowDataLabelsOverMaximum"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2504) //= 0x9c8. The runtime will prefer the VTID if present
  @VTID(170)
  boolean getShowDataLabelsOverMaximum();


  /**
   * <p>
   * Setter method for the COM property "ShowDataLabelsOverMaximum"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2504) //= 0x9c8. The runtime will prefer the VTID if present
  @VTID(171)
  void setShowDataLabelsOverMaximum(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "SideWall"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Walls
   */

  @DISPID(2505) //= 0x9c9. The runtime will prefer the VTID if present
  @VTID(172)
  com.exceljava.com4j.excel.Walls getSideWall();


  /**
   * <p>
   * Getter method for the COM property "BackWall"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Walls
   */

  @DISPID(2506) //= 0x9ca. The runtime will prefer the VTID if present
  @VTID(173)
  com.exceljava.com4j.excel.Walls getBackWall();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(174)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void printOut();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(174)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(174)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(174)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(174)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(174)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, printToFile, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(174)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, printToFile, collate, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(174)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, printToFile, collate, prToFileName, 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @param prToFileName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(174)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object prToFileName);

  /**
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @param prToFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(174)
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object prToFileName,
    @LCID int lcid);


  /**
   * @param filename Mandatory java.lang.String parameter.
   */

  @DISPID(2507) //= 0x9cb. The runtime will prefer the VTID if present
  @VTID(175)
  void applyChartTemplate(
    java.lang.String filename);


  /**
   * @param filename Mandatory java.lang.String parameter.
   */

  @DISPID(2508) //= 0x9cc. The runtime will prefer the VTID if present
  @VTID(176)
  void saveChartTemplate(
    java.lang.String filename);


  /**
   * @param name Mandatory java.lang.Object parameter.
   */

  @DISPID(219) //= 0xdb. The runtime will prefer the VTID if present
  @VTID(177)
  void setDefaultChart(
    @MarshalAs(NativeType.VARIANT) java.lang.Object name);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter filename is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter quality is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter includeDocProperties is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ExportAsFixedFormat(type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   */

  @DISPID(2493) //= 0x9bd. The runtime will prefer the VTID if present
  @VTID(178)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter quality is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter includeDocProperties is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ExportAsFixedFormat(type, filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2493) //= 0x9bd. The runtime will prefer the VTID if present
  @VTID(178)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter includeDocProperties is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ExportAsFixedFormat(type, filename, quality, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2493) //= 0x9bd. The runtime will prefer the VTID if present
  @VTID(178)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ExportAsFixedFormat(type, filename, quality, includeDocProperties, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2493) //= 0x9bd. The runtime will prefer the VTID if present
  @VTID(178)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ExportAsFixedFormat(type, filename, quality, includeDocProperties, ignorePrintAreas, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2493) //= 0x9bd. The runtime will prefer the VTID if present
  @VTID(178)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ExportAsFixedFormat(type, filename, quality, includeDocProperties, ignorePrintAreas, from, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2493) //= 0x9bd. The runtime will prefer the VTID if present
  @VTID(178)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ExportAsFixedFormat(type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2493) //= 0x9bd. The runtime will prefer the VTID if present
  @VTID(178)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _ExportAsFixedFormat(type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param openAfterPublish Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2493) //= 0x9bd. The runtime will prefer the VTID if present
  @VTID(178)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object openAfterPublish);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param openAfterPublish Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fixedFormatExtClassPtr Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2493) //= 0x9bd. The runtime will prefer the VTID if present
  @VTID(178)
  void _ExportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object openAfterPublish,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fixedFormatExtClassPtr);


  /**
   * <p>
   * Getter method for the COM property "ChartStyle"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(2509) //= 0x9cd. The runtime will prefer the VTID if present
  @VTID(179)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getChartStyle();


  /**
   * <p>
   * Setter method for the COM property "ChartStyle"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(2509) //= 0x9cd. The runtime will prefer the VTID if present
  @VTID(180)
  void setChartStyle(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   */

  @DISPID(2510) //= 0x9ce. The runtime will prefer the VTID if present
  @VTID(181)
  void clearToMatchStyle();


  /**
   * <p>
   * Getter method for the COM property "PrintedCommentPages"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(2857) //= 0xb29. The runtime will prefer the VTID if present
  @VTID(182)
  int getPrintedCommentPages();


  /**
   * <p>
   * Getter method for the COM property "Dummy24"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2858) //= 0xb2a. The runtime will prefer the VTID if present
  @VTID(183)
  boolean getDummy24();


  /**
   * <p>
   * Setter method for the COM property "Dummy24"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2858) //= 0xb2a. The runtime will prefer the VTID if present
  @VTID(184)
  void setDummy24(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Dummy25"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2859) //= 0xb2b. The runtime will prefer the VTID if present
  @VTID(185)
  boolean getDummy25();


  /**
   * <p>
   * Setter method for the COM property "Dummy25"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2859) //= 0xb2b. The runtime will prefer the VTID if present
  @VTID(186)
  void setDummy25(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowReportFilterFieldButtons"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2860) //= 0xb2c. The runtime will prefer the VTID if present
  @VTID(187)
  boolean getShowReportFilterFieldButtons();


  /**
   * <p>
   * Setter method for the COM property "ShowReportFilterFieldButtons"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2860) //= 0xb2c. The runtime will prefer the VTID if present
  @VTID(188)
  void setShowReportFilterFieldButtons(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowLegendFieldButtons"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2861) //= 0xb2d. The runtime will prefer the VTID if present
  @VTID(189)
  boolean getShowLegendFieldButtons();


  /**
   * <p>
   * Setter method for the COM property "ShowLegendFieldButtons"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2861) //= 0xb2d. The runtime will prefer the VTID if present
  @VTID(190)
  void setShowLegendFieldButtons(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowAxisFieldButtons"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2862) //= 0xb2e. The runtime will prefer the VTID if present
  @VTID(191)
  boolean getShowAxisFieldButtons();


  /**
   * <p>
   * Setter method for the COM property "ShowAxisFieldButtons"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2862) //= 0xb2e. The runtime will prefer the VTID if present
  @VTID(192)
  void setShowAxisFieldButtons(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowValueFieldButtons"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2863) //= 0xb2f. The runtime will prefer the VTID if present
  @VTID(193)
  boolean getShowValueFieldButtons();


  /**
   * <p>
   * Setter method for the COM property "ShowValueFieldButtons"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2863) //= 0xb2f. The runtime will prefer the VTID if present
  @VTID(194)
  void setShowValueFieldButtons(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowAllFieldButtons"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2864) //= 0xb30. The runtime will prefer the VTID if present
  @VTID(195)
  boolean getShowAllFieldButtons();


  /**
   * <p>
   * Setter method for the COM property "ShowAllFieldButtons"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2864) //= 0xb30. The runtime will prefer the VTID if present
  @VTID(196)
  void setShowAllFieldButtons(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * fullSeriesCollection(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(3047) //= 0xbe7. The runtime will prefer the VTID if present
  @VTID(197)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject fullSeriesCollection();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * fullSeriesCollection(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(3047) //= 0xbe7. The runtime will prefer the VTID if present
  @VTID(197)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject fullSeriesCollection(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(3047) //= 0xbe7. The runtime will prefer the VTID if present
  @VTID(197)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject fullSeriesCollection(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "CategoryLabelLevel"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCategoryLabelLevel
   */

  @DISPID(3048) //= 0xbe8. The runtime will prefer the VTID if present
  @VTID(198)
  com.exceljava.com4j.excel.XlCategoryLabelLevel getCategoryLabelLevel();


  /**
   * <p>
   * Setter method for the COM property "CategoryLabelLevel"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlCategoryLabelLevel parameter.
   */

  @DISPID(3048) //= 0xbe8. The runtime will prefer the VTID if present
  @VTID(199)
  void setCategoryLabelLevel(
    com.exceljava.com4j.excel.XlCategoryLabelLevel rhs);


  /**
   * <p>
   * Getter method for the COM property "SeriesNameLevel"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSeriesNameLevel
   */

  @DISPID(3049) //= 0xbe9. The runtime will prefer the VTID if present
  @VTID(200)
  com.exceljava.com4j.excel.XlSeriesNameLevel getSeriesNameLevel();


  /**
   * <p>
   * Setter method for the COM property "SeriesNameLevel"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSeriesNameLevel parameter.
   */

  @DISPID(3049) //= 0xbe9. The runtime will prefer the VTID if present
  @VTID(201)
  void setSeriesNameLevel(
    com.exceljava.com4j.excel.XlSeriesNameLevel rhs);


  /**
   * <p>
   * Getter method for the COM property "HasHiddenContent"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(3050) //= 0xbea. The runtime will prefer the VTID if present
  @VTID(202)
  boolean getHasHiddenContent();


  /**
   */

  @DISPID(3051) //= 0xbeb. The runtime will prefer the VTID if present
  @VTID(203)
  void deleteHiddenContent();


  /**
   * <p>
   * Getter method for the COM property "ChartColor"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(3052) //= 0xbec. The runtime will prefer the VTID if present
  @VTID(204)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getChartColor();


  /**
   * <p>
   * Setter method for the COM property "ChartColor"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(3052) //= 0xbec. The runtime will prefer the VTID if present
  @VTID(205)
  void setChartColor(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   */

  @DISPID(3053) //= 0xbed. The runtime will prefer the VTID if present
  @VTID(206)
  void clearToMatchColorStyle();


  /**
   * <p>
   * Getter method for the COM property "ShowExpandCollapseEntireFieldButtons"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(3166) //= 0xc5e. The runtime will prefer the VTID if present
  @VTID(207)
  boolean getShowExpandCollapseEntireFieldButtons();


  /**
   * <p>
   * Setter method for the COM property "ShowExpandCollapseEntireFieldButtons"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(3166) //= 0xc5e. The runtime will prefer the VTID if present
  @VTID(208)
  void setShowExpandCollapseEntireFieldButtons(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter filename is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter quality is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter includeDocProperties is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exportAsFixedFormat(type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(209)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter quality is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter includeDocProperties is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exportAsFixedFormat(type, filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(209)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter includeDocProperties is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exportAsFixedFormat(type, filename, quality, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(209)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exportAsFixedFormat(type, filename, quality, includeDocProperties, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(209)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exportAsFixedFormat(type, filename, quality, includeDocProperties, ignorePrintAreas, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(209)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exportAsFixedFormat(type, filename, quality, includeDocProperties, ignorePrintAreas, from, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(209)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter openAfterPublish is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exportAsFixedFormat(type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(209)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter fixedFormatExtClassPtr is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exportAsFixedFormat(type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param openAfterPublish Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(209)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object openAfterPublish);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter workIdentity is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * exportAsFixedFormat(type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish, fixedFormatExtClassPtr, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param openAfterPublish Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fixedFormatExtClassPtr Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(209)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object openAfterPublish,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fixedFormatExtClassPtr);

  /**
   * @param type Mandatory com.exceljava.com4j.excel.XlFixedFormatType parameter.
   * @param filename Optional parameter. Default value is com4j.Variant.getMissing()
   * @param quality Optional parameter. Default value is com4j.Variant.getMissing()
   * @param includeDocProperties Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param openAfterPublish Optional parameter. Default value is com4j.Variant.getMissing()
   * @param fixedFormatExtClassPtr Optional parameter. Default value is com4j.Variant.getMissing()
   * @param workIdentity Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3175) //= 0xc67. The runtime will prefer the VTID if present
  @VTID(209)
  void exportAsFixedFormat(
    com.exceljava.com4j.excel.XlFixedFormatType type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object quality,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object includeDocProperties,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object openAfterPublish,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fixedFormatExtClassPtr,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object workIdentity);


  /**
   * @param id Mandatory java.lang.String parameter.
   * @param value Mandatory java.lang.Object parameter.
   */

  @DISPID(3253) //= 0xcb5. The runtime will prefer the VTID if present
  @VTID(210)
  void setProperty(
    java.lang.String id,
    @MarshalAs(NativeType.VARIANT) java.lang.Object value);


  /**
   * @param id Mandatory java.lang.String parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(3256) //= 0xcb8. The runtime will prefer the VTID if present
  @VTID(211)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getProperty(
    java.lang.String id);


  /**
   * <p>
   * Getter method for the COM property "DisplayValueNotAvailableAsBlank"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(3278) //= 0xcce. The runtime will prefer the VTID if present
  @VTID(212)
  boolean getDisplayValueNotAvailableAsBlank();


  /**
   * <p>
   * Setter method for the COM property "DisplayValueNotAvailableAsBlank"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(3278) //= 0xcce. The runtime will prefer the VTID if present
  @VTID(213)
  void setDisplayValueNotAvailableAsBlank(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter fileFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(214)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void saveAs(
    java.lang.String filename);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(214)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void saveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter writeResPassword is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, password, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(214)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void saveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter readOnlyRecommended is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, password, writeResPassword, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(214)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void saveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter createBackup is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(214)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  void saveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter addToMru is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(214)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void saveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter textCodepage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(214)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void saveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter textVisualLayout is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(214)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void saveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter local is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * saveAs(filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage, textVisualLayout, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(214)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void saveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout);

  /**
   * @param filename Mandatory java.lang.String parameter.
   * @param fileFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param writeResPassword Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readOnlyRecommended Optional parameter. Default value is com4j.Variant.getMissing()
   * @param createBackup Optional parameter. Default value is com4j.Variant.getMissing()
   * @param addToMru Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textCodepage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param textVisualLayout Optional parameter. Default value is com4j.Variant.getMissing()
   * @param local Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(3174) //= 0xc66. The runtime will prefer the VTID if present
  @VTID(214)
  void saveAs(
    java.lang.String filename,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object fileFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object writeResPassword,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readOnlyRecommended,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object createBackup,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object addToMru,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textCodepage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object textVisualLayout,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object local);


  // Properties:
}
