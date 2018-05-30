package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000208D8-0000-0000-C000-000000000046}")
public interface _Worksheet extends Com4jObject {
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
   * <p>
   * Getter method for the COM property "ProtectScenarios"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getProtectScenarios(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(294) //= 0x126. The runtime will prefer the VTID if present
  @VTID(35)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getProtectScenarios();

  /**
   * <p>
   * Getter method for the COM property "ProtectScenarios"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(294) //= 0x126. The runtime will prefer the VTID if present
  @VTID(35)
  boolean getProtectScenarios(
    @LCID int lcid);


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
   * Getter method for the COM property "TransitionExpEval"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getTransitionExpEval(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(401) //= 0x191. The runtime will prefer the VTID if present
  @VTID(42)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getTransitionExpEval();

  /**
   * <p>
   * Getter method for the COM property "TransitionExpEval"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(401) //= 0x191. The runtime will prefer the VTID if present
  @VTID(42)
  boolean getTransitionExpEval(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "TransitionExpEval"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setTransitionExpEval(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(401) //= 0x191. The runtime will prefer the VTID if present
  @VTID(43)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setTransitionExpEval(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "TransitionExpEval"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(401) //= 0x191. The runtime will prefer the VTID if present
  @VTID(43)
  void setTransitionExpEval(
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
   * arcs(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(760) //= 0x2f8. The runtime will prefer the VTID if present
  @VTID(44)
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
  @VTID(44)
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
  @VTID(44)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject arcs(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "AutoFilterMode"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAutoFilterMode(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(792) //= 0x318. The runtime will prefer the VTID if present
  @VTID(45)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getAutoFilterMode();

  /**
   * <p>
   * Getter method for the COM property "AutoFilterMode"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(792) //= 0x318. The runtime will prefer the VTID if present
  @VTID(45)
  boolean getAutoFilterMode(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "AutoFilterMode"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setAutoFilterMode(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(792) //= 0x318. The runtime will prefer the VTID if present
  @VTID(46)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setAutoFilterMode(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "AutoFilterMode"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(792) //= 0x318. The runtime will prefer the VTID if present
  @VTID(46)
  void setAutoFilterMode(
    @LCID int lcid,
    boolean rhs);


  /**
   * @param filename Mandatory java.lang.String parameter.
   */

  @DISPID(1188) //= 0x4a4. The runtime will prefer the VTID if present
  @VTID(47)
  void setBackgroundPicture(
    java.lang.String filename);


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
  @VTID(48)
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
  @VTID(48)
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
  @VTID(48)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject buttons(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
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
   * calculate(1033);
   * </code>
   * </p>
   */

  @DISPID(279) //= 0x117. The runtime will prefer the VTID if present
  @VTID(49)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void calculate();

  /**
   * @param lcid Mandatory int parameter.
   */

  @DISPID(279) //= 0x117. The runtime will prefer the VTID if present
  @VTID(49)
  void calculate(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "EnableCalculation"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1424) //= 0x590. The runtime will prefer the VTID if present
  @VTID(50)
  boolean getEnableCalculation();


  /**
   * <p>
   * Setter method for the COM property "EnableCalculation"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1424) //= 0x590. The runtime will prefer the VTID if present
  @VTID(51)
  void setEnableCalculation(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Cells"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @DISPID(238) //= 0xee. The runtime will prefer the VTID if present
  @VTID(52)
  com.exceljava.com4j.excel.Range getCells();


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
  @VTID(53)
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
  @VTID(53)
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
  @VTID(53)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject chartObjects(
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
   * checkBoxes(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(824) //= 0x338. The runtime will prefer the VTID if present
  @VTID(54)
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
  @VTID(54)
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
  @VTID(54)
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
  @VTID(55)
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
  @VTID(55)
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
  @VTID(55)
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
  @VTID(55)
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
  @VTID(55)
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
  @VTID(55)
  void checkSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreUppercase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object alwaysSuggest,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object spellLang,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "CircularReference"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getCircularReference(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @DISPID(1069) //= 0x42d. The runtime will prefer the VTID if present
  @VTID(56)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.Range getCircularReference();

  /**
   * <p>
   * Getter method for the COM property "CircularReference"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @DISPID(1069) //= 0x42d. The runtime will prefer the VTID if present
  @VTID(56)
  com.exceljava.com4j.excel.Range getCircularReference(
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
   * clearArrows(1033);
   * </code>
   * </p>
   */

  @DISPID(970) //= 0x3ca. The runtime will prefer the VTID if present
  @VTID(57)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void clearArrows();

  /**
   * @param lcid Mandatory int parameter.
   */

  @DISPID(970) //= 0x3ca. The runtime will prefer the VTID if present
  @VTID(57)
  void clearArrows(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Columns"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @DISPID(241) //= 0xf1. The runtime will prefer the VTID if present
  @VTID(58)
  com.exceljava.com4j.excel.Range getColumns();


  /**
   * <p>
   * Getter method for the COM property "ConsolidationFunction"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getConsolidationFunction(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlConsolidationFunction
   */

  @DISPID(789) //= 0x315. The runtime will prefer the VTID if present
  @VTID(59)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.XlConsolidationFunction getConsolidationFunction();

  /**
   * <p>
   * Getter method for the COM property "ConsolidationFunction"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.XlConsolidationFunction
   */

  @DISPID(789) //= 0x315. The runtime will prefer the VTID if present
  @VTID(59)
  com.exceljava.com4j.excel.XlConsolidationFunction getConsolidationFunction(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "ConsolidationOptions"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getConsolidationOptions(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(790) //= 0x316. The runtime will prefer the VTID if present
  @VTID(60)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getConsolidationOptions();

  /**
   * <p>
   * Getter method for the COM property "ConsolidationOptions"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(790) //= 0x316. The runtime will prefer the VTID if present
  @VTID(60)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getConsolidationOptions(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "ConsolidationSources"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getConsolidationSources(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(791) //= 0x317. The runtime will prefer the VTID if present
  @VTID(61)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getConsolidationSources();

  /**
   * <p>
   * Getter method for the COM property "ConsolidationSources"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(791) //= 0x317. The runtime will prefer the VTID if present
  @VTID(61)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getConsolidationSources(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "DisplayAutomaticPageBreaks"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getDisplayAutomaticPageBreaks(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(643) //= 0x283. The runtime will prefer the VTID if present
  @VTID(62)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getDisplayAutomaticPageBreaks();

  /**
   * <p>
   * Getter method for the COM property "DisplayAutomaticPageBreaks"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(643) //= 0x283. The runtime will prefer the VTID if present
  @VTID(62)
  boolean getDisplayAutomaticPageBreaks(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "DisplayAutomaticPageBreaks"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setDisplayAutomaticPageBreaks(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(643) //= 0x283. The runtime will prefer the VTID if present
  @VTID(63)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setDisplayAutomaticPageBreaks(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "DisplayAutomaticPageBreaks"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(643) //= 0x283. The runtime will prefer the VTID if present
  @VTID(63)
  void setDisplayAutomaticPageBreaks(
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
   * drawings(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(772) //= 0x304. The runtime will prefer the VTID if present
  @VTID(64)
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
  @VTID(64)
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
  @VTID(64)
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
  @VTID(65)
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
  @VTID(65)
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
  @VTID(65)
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
  @VTID(66)
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
  @VTID(66)
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
  @VTID(66)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject dropDowns(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "EnableAutoFilter"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getEnableAutoFilter(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1156) //= 0x484. The runtime will prefer the VTID if present
  @VTID(67)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getEnableAutoFilter();

  /**
   * <p>
   * Getter method for the COM property "EnableAutoFilter"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(1156) //= 0x484. The runtime will prefer the VTID if present
  @VTID(67)
  boolean getEnableAutoFilter(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "EnableAutoFilter"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setEnableAutoFilter(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1156) //= 0x484. The runtime will prefer the VTID if present
  @VTID(68)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setEnableAutoFilter(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "EnableAutoFilter"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1156) //= 0x484. The runtime will prefer the VTID if present
  @VTID(68)
  void setEnableAutoFilter(
    @LCID int lcid,
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "EnableSelection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlEnableSelection
   */

  @DISPID(1425) //= 0x591. The runtime will prefer the VTID if present
  @VTID(69)
  com.exceljava.com4j.excel.XlEnableSelection getEnableSelection();


  /**
   * <p>
   * Setter method for the COM property "EnableSelection"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlEnableSelection parameter.
   */

  @DISPID(1425) //= 0x591. The runtime will prefer the VTID if present
  @VTID(70)
  void setEnableSelection(
    com.exceljava.com4j.excel.XlEnableSelection rhs);


  /**
   * <p>
   * Getter method for the COM property "EnableOutlining"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getEnableOutlining(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1157) //= 0x485. The runtime will prefer the VTID if present
  @VTID(71)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getEnableOutlining();

  /**
   * <p>
   * Getter method for the COM property "EnableOutlining"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(1157) //= 0x485. The runtime will prefer the VTID if present
  @VTID(71)
  boolean getEnableOutlining(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "EnableOutlining"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setEnableOutlining(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1157) //= 0x485. The runtime will prefer the VTID if present
  @VTID(72)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setEnableOutlining(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "EnableOutlining"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1157) //= 0x485. The runtime will prefer the VTID if present
  @VTID(72)
  void setEnableOutlining(
    @LCID int lcid,
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "EnablePivotTable"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getEnablePivotTable(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1158) //= 0x486. The runtime will prefer the VTID if present
  @VTID(73)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getEnablePivotTable();

  /**
   * <p>
   * Getter method for the COM property "EnablePivotTable"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(1158) //= 0x486. The runtime will prefer the VTID if present
  @VTID(73)
  boolean getEnablePivotTable(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "EnablePivotTable"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setEnablePivotTable(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1158) //= 0x486. The runtime will prefer the VTID if present
  @VTID(74)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setEnablePivotTable(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "EnablePivotTable"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1158) //= 0x486. The runtime will prefer the VTID if present
  @VTID(74)
  void setEnablePivotTable(
    @LCID int lcid,
    boolean rhs);


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
  @VTID(75)
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
  @VTID(75)
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
  @VTID(76)
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
  @VTID(76)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object _Evaluate(
    @MarshalAs(NativeType.VARIANT) java.lang.Object name,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "FilterMode"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getFilterMode(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(800) //= 0x320. The runtime will prefer the VTID if present
  @VTID(77)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getFilterMode();

  /**
   * <p>
   * Getter method for the COM property "FilterMode"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(800) //= 0x320. The runtime will prefer the VTID if present
  @VTID(77)
  boolean getFilterMode(
    @LCID int lcid);


  /**
   */

  @DISPID(1426) //= 0x592. The runtime will prefer the VTID if present
  @VTID(78)
  void resetAllPageBreaks();


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
  @VTID(79)
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
  @VTID(79)
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
  @VTID(79)
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
  @VTID(80)
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
  @VTID(80)
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
  @VTID(80)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject groupObjects(
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
   * labels(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(841) //= 0x349. The runtime will prefer the VTID if present
  @VTID(81)
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
  @VTID(81)
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
  @VTID(81)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject labels(
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
  @VTID(82)
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
  @VTID(82)
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
  @VTID(82)
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
  @VTID(83)
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
  @VTID(83)
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
  @VTID(83)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject listBoxes(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Names"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Names
   */

  @DISPID(442) //= 0x1ba. The runtime will prefer the VTID if present
  @VTID(84)
  com.exceljava.com4j.excel.Names getNames();


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
  @VTID(85)
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
  @VTID(85)
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
  @VTID(85)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject oleObjects(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "OnCalculate"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getOnCalculate(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(625) //= 0x271. The runtime will prefer the VTID if present
  @VTID(86)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getOnCalculate();

  /**
   * <p>
   * Getter method for the COM property "OnCalculate"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(625) //= 0x271. The runtime will prefer the VTID if present
  @VTID(86)
  java.lang.String getOnCalculate(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "OnCalculate"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setOnCalculate(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(625) //= 0x271. The runtime will prefer the VTID if present
  @VTID(87)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setOnCalculate(
    java.lang.String rhs);

  /**
   * <p>
   * Setter method for the COM property "OnCalculate"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(625) //= 0x271. The runtime will prefer the VTID if present
  @VTID(87)
  void setOnCalculate(
    @LCID int lcid,
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "OnData"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getOnData(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(629) //= 0x275. The runtime will prefer the VTID if present
  @VTID(88)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getOnData();

  /**
   * <p>
   * Getter method for the COM property "OnData"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(629) //= 0x275. The runtime will prefer the VTID if present
  @VTID(88)
  java.lang.String getOnData(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "OnData"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setOnData(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(629) //= 0x275. The runtime will prefer the VTID if present
  @VTID(89)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setOnData(
    java.lang.String rhs);

  /**
   * <p>
   * Setter method for the COM property "OnData"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(629) //= 0x275. The runtime will prefer the VTID if present
  @VTID(89)
  void setOnData(
    @LCID int lcid,
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "OnEntry"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getOnEntry(1033);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(627) //= 0x273. The runtime will prefer the VTID if present
  @VTID(90)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  java.lang.String getOnEntry();

  /**
   * <p>
   * Getter method for the COM property "OnEntry"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(627) //= 0x273. The runtime will prefer the VTID if present
  @VTID(90)
  java.lang.String getOnEntry(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "OnEntry"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setOnEntry(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(627) //= 0x273. The runtime will prefer the VTID if present
  @VTID(91)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setOnEntry(
    java.lang.String rhs);

  /**
   * <p>
   * Setter method for the COM property "OnEntry"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(627) //= 0x273. The runtime will prefer the VTID if present
  @VTID(91)
  void setOnEntry(
    @LCID int lcid,
    java.lang.String rhs);


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
  @VTID(92)
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
  @VTID(92)
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
  @VTID(92)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject optionButtons(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Outline"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Outline
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(93)
  com.exceljava.com4j.excel.Outline getOutline();


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
  @VTID(94)
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
  @VTID(94)
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
  @VTID(94)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject ovals(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter destination is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter link is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * paste(com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(211) //= 0xd3. The runtime will prefer the VTID if present
  @VTID(95)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void paste();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter link is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * paste(destination, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(211) //= 0xd3. The runtime will prefer the VTID if present
  @VTID(95)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void paste(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * paste(destination, link, 1033);
   * </code>
   * </p>
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(211) //= 0xd3. The runtime will prefer the VTID if present
  @VTID(95)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void paste(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link);

  /**
   * @param destination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(211) //= 0xd3. The runtime will prefer the VTID if present
  @VTID(95)
  void paste(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object destination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter link is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter displayAsIcon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PasteSpecial(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(1027) //= 0x403. The runtime will prefer the VTID if present
  @VTID(96)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _PasteSpecial();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter link is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter displayAsIcon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PasteSpecial(format, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1027) //= 0x403. The runtime will prefer the VTID if present
  @VTID(96)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _PasteSpecial(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter displayAsIcon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PasteSpecial(format, link, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1027) //= 0x403. The runtime will prefer the VTID if present
  @VTID(96)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void _PasteSpecial(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter iconFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PasteSpecial(format, link, displayAsIcon, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1027) //= 0x403. The runtime will prefer the VTID if present
  @VTID(96)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void _PasteSpecial(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayAsIcon);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter iconIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PasteSpecial(format, link, displayAsIcon, iconFileName, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconFileName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1027) //= 0x403. The runtime will prefer the VTID if present
  @VTID(96)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void _PasteSpecial(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayAsIcon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconFileName);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PasteSpecial(format, link, displayAsIcon, iconFileName, iconIndex, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconIndex Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1027) //= 0x403. The runtime will prefer the VTID if present
  @VTID(96)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void _PasteSpecial(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayAsIcon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconIndex);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _PasteSpecial(format, link, displayAsIcon, iconFileName, iconIndex, iconLabel, 1033);
   * </code>
   * </p>
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconLabel Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1027) //= 0x403. The runtime will prefer the VTID if present
  @VTID(96)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void _PasteSpecial(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayAsIcon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconLabel);

  /**
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconLabel Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(1027) //= 0x403. The runtime will prefer the VTID if present
  @VTID(96)
  void _PasteSpecial(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayAsIcon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconLabel,
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
   * pictures(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(771) //= 0x303. The runtime will prefer the VTID if present
  @VTID(97)
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
  @VTID(97)
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
  @VTID(97)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject pictures(
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
   * pivotTables(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(690) //= 0x2b2. The runtime will prefer the VTID if present
  @VTID(98)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject pivotTables();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTables(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(690) //= 0x2b2. The runtime will prefer the VTID if present
  @VTID(98)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject pivotTables(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(690) //= 0x2b2. The runtime will prefer the VTID if present
  @VTID(98)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject pivotTables(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter sourceType is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter sourceData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tableDestination is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tableName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter rowGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter saveData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasAutoFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoPage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(99)
  @UseDefaultValues(paramIndexMapping = {17}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=17)
  com.exceljava.com4j.excel.PivotTable pivotTableWizard();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter sourceData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tableDestination is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tableName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter rowGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter saveData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasAutoFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoPage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(99)
  @UseDefaultValues(paramIndexMapping = {0, 17}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=17)
  com.exceljava.com4j.excel.PivotTable pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter tableDestination is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tableName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter rowGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter saveData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasAutoFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoPage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(99)
  @UseDefaultValues(paramIndexMapping = {0, 1, 17}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=17)
  com.exceljava.com4j.excel.PivotTable pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter tableName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter rowGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter saveData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasAutoFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoPage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(99)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 17}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=17)
  com.exceljava.com4j.excel.PivotTable pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter rowGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter columnGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter saveData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasAutoFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoPage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(99)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 17}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=17)
  com.exceljava.com4j.excel.PivotTable pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter columnGrand is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter saveData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasAutoFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoPage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(99)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 17}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=17)
  com.exceljava.com4j.excel.PivotTable pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter saveData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter hasAutoFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoPage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(99)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 17}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=17)
  com.exceljava.com4j.excel.PivotTable pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter hasAutoFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter autoPage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(99)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 17}, optParamIndex = {7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=17)
  com.exceljava.com4j.excel.PivotTable pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter autoPage is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(99)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 17}, optParamIndex = {8, 9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=17)
  com.exceljava.com4j.excel.PivotTable pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter reserved is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoPage Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(99)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 17}, optParamIndex = {9, 10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=17)
  com.exceljava.com4j.excel.PivotTable pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoPage);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter backgroundQuery is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoPage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reserved Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(99)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 17}, optParamIndex = {10, 11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=17)
  com.exceljava.com4j.excel.PivotTable pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoPage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reserved);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter optimizeCache is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoPage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reserved Optional parameter. Default value is com4j.Variant.getMissing()
   * @param backgroundQuery Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(99)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 17}, optParamIndex = {11, 12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=17)
  com.exceljava.com4j.excel.PivotTable pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoPage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reserved,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object backgroundQuery);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter pageFieldOrder is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoPage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reserved Optional parameter. Default value is com4j.Variant.getMissing()
   * @param backgroundQuery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param optimizeCache Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(99)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 17}, optParamIndex = {12, 13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=17)
  com.exceljava.com4j.excel.PivotTable pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoPage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reserved,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object backgroundQuery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object optimizeCache);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter pageFieldWrapCount is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoPage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reserved Optional parameter. Default value is com4j.Variant.getMissing()
   * @param backgroundQuery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param optimizeCache Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFieldOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(99)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 17}, optParamIndex = {13, 14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  @ReturnValue(index=17)
  com.exceljava.com4j.excel.PivotTable pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoPage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reserved,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object backgroundQuery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object optimizeCache,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageFieldOrder);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter readData is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoPage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reserved Optional parameter. Default value is com4j.Variant.getMissing()
   * @param backgroundQuery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param optimizeCache Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFieldOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFieldWrapCount Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(99)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 17}, optParamIndex = {14, 15, 16}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  @ReturnValue(index=17)
  com.exceljava.com4j.excel.PivotTable pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoPage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reserved,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object backgroundQuery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object optimizeCache,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageFieldOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageFieldWrapCount);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter connection is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount, readData, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoPage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reserved Optional parameter. Default value is com4j.Variant.getMissing()
   * @param backgroundQuery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param optimizeCache Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFieldOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFieldWrapCount Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readData Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(99)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 17}, optParamIndex = {15, 16}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(index=17)
  com.exceljava.com4j.excel.PivotTable pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoPage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reserved,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object backgroundQuery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object optimizeCache,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageFieldOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageFieldWrapCount,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readData);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pivotTableWizard(sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount, readData, connection, 1033);
   * </code>
   * </p>
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoPage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reserved Optional parameter. Default value is com4j.Variant.getMissing()
   * @param backgroundQuery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param optimizeCache Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFieldOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFieldWrapCount Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param connection Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(99)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 17}, optParamIndex = {16}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=17)
  com.exceljava.com4j.excel.PivotTable pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoPage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reserved,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object backgroundQuery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object optimizeCache,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageFieldOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageFieldWrapCount,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object connection);

  /**
   * @param sourceType Optional parameter. Default value is com4j.Variant.getMissing()
   * @param sourceData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableDestination Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tableName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param rowGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param columnGrand Optional parameter. Default value is com4j.Variant.getMissing()
   * @param saveData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param hasAutoFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param autoPage Optional parameter. Default value is com4j.Variant.getMissing()
   * @param reserved Optional parameter. Default value is com4j.Variant.getMissing()
   * @param backgroundQuery Optional parameter. Default value is com4j.Variant.getMissing()
   * @param optimizeCache Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFieldOrder Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pageFieldWrapCount Optional parameter. Default value is com4j.Variant.getMissing()
   * @param readData Optional parameter. Default value is com4j.Variant.getMissing()
   * @param connection Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @DISPID(684) //= 0x2ac. The runtime will prefer the VTID if present
  @VTID(99)
  com.exceljava.com4j.excel.PivotTable pivotTableWizard(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceType,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object sourceData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableDestination,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tableName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object rowGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object columnGrand,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object saveData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object hasAutoFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object autoPage,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object reserved,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object backgroundQuery,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object optimizeCache,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageFieldOrder,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pageFieldWrapCount,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object readData,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object connection,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Range"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter cell2 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getRange(cell1, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param cell1 Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @DISPID(197) //= 0xc5. The runtime will prefer the VTID if present
  @VTID(100)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.Range getRange(
    @MarshalAs(NativeType.VARIANT) java.lang.Object cell1);

  /**
   * <p>
   * Getter method for the COM property "Range"
   * </p>
   * @param cell1 Mandatory java.lang.Object parameter.
   * @param cell2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @DISPID(197) //= 0xc5. The runtime will prefer the VTID if present
  @VTID(100)
  com.exceljava.com4j.excel.Range getRange(
    @MarshalAs(NativeType.VARIANT) java.lang.Object cell1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object cell2);


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
  @VTID(101)
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
  @VTID(101)
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
  @VTID(101)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject rectangles(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Rows"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @DISPID(258) //= 0x102. The runtime will prefer the VTID if present
  @VTID(102)
  com.exceljava.com4j.excel.Range getRows();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * scenarios(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(908) //= 0x38c. The runtime will prefer the VTID if present
  @VTID(103)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject scenarios();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * scenarios(index, 1033);
   * </code>
   * </p>
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(908) //= 0x38c. The runtime will prefer the VTID if present
  @VTID(103)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(type=NativeType.Dispatch,index=2)
  com4j.Com4jObject scenarios(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(908) //= 0x38c. The runtime will prefer the VTID if present
  @VTID(103)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject scenarios(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "ScrollArea"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1433) //= 0x599. The runtime will prefer the VTID if present
  @VTID(104)
  java.lang.String getScrollArea();


  /**
   * <p>
   * Setter method for the COM property "ScrollArea"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(1433) //= 0x599. The runtime will prefer the VTID if present
  @VTID(105)
  void setScrollArea(
    java.lang.String rhs);


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
  @VTID(106)
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
  @VTID(106)
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
  @VTID(106)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject scrollBars(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
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
   * showAllData(1033);
   * </code>
   * </p>
   */

  @DISPID(794) //= 0x31a. The runtime will prefer the VTID if present
  @VTID(107)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void showAllData();

  /**
   * @param lcid Mandatory int parameter.
   */

  @DISPID(794) //= 0x31a. The runtime will prefer the VTID if present
  @VTID(107)
  void showAllData(
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
   * showDataForm(1033);
   * </code>
   * </p>
   */

  @DISPID(409) //= 0x199. The runtime will prefer the VTID if present
  @VTID(108)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void showDataForm();

  /**
   * @param lcid Mandatory int parameter.
   */

  @DISPID(409) //= 0x199. The runtime will prefer the VTID if present
  @VTID(108)
  void showDataForm(
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
   * spinners(com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(838) //= 0x346. The runtime will prefer the VTID if present
  @VTID(109)
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
  @VTID(109)
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
  @VTID(109)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject spinners(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "StandardHeight"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getStandardHeight(1033);
   * </code>
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(407) //= 0x197. The runtime will prefer the VTID if present
  @VTID(110)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  double getStandardHeight();

  /**
   * <p>
   * Getter method for the COM property "StandardHeight"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type double
   */

  @DISPID(407) //= 0x197. The runtime will prefer the VTID if present
  @VTID(110)
  double getStandardHeight(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "StandardWidth"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getStandardWidth(1033);
   * </code>
   * </p>
   * @return  Returns a value of type double
   */

  @DISPID(408) //= 0x198. The runtime will prefer the VTID if present
  @VTID(111)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  double getStandardWidth();

  /**
   * <p>
   * Getter method for the COM property "StandardWidth"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type double
   */

  @DISPID(408) //= 0x198. The runtime will prefer the VTID if present
  @VTID(111)
  double getStandardWidth(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "StandardWidth"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setStandardWidth(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(408) //= 0x198. The runtime will prefer the VTID if present
  @VTID(112)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setStandardWidth(
    double rhs);

  /**
   * <p>
   * Setter method for the COM property "StandardWidth"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory double parameter.
   */

  @DISPID(408) //= 0x198. The runtime will prefer the VTID if present
  @VTID(112)
  void setStandardWidth(
    @LCID int lcid,
    double rhs);


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
  @VTID(113)
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
  @VTID(113)
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
  @VTID(113)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject textBoxes(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "TransitionFormEntry"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getTransitionFormEntry(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(402) //= 0x192. The runtime will prefer the VTID if present
  @VTID(114)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getTransitionFormEntry();

  /**
   * <p>
   * Getter method for the COM property "TransitionFormEntry"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(402) //= 0x192. The runtime will prefer the VTID if present
  @VTID(114)
  boolean getTransitionFormEntry(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "TransitionFormEntry"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setTransitionFormEntry(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(402) //= 0x192. The runtime will prefer the VTID if present
  @VTID(115)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setTransitionFormEntry(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "TransitionFormEntry"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(402) //= 0x192. The runtime will prefer the VTID if present
  @VTID(115)
  void setTransitionFormEntry(
    @LCID int lcid,
    boolean rhs);


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
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSheetType
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(116)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.XlSheetType getType();

  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSheetType
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(116)
  com.exceljava.com4j.excel.XlSheetType getType(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "UsedRange"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getUsedRange(1033);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @DISPID(412) //= 0x19c. The runtime will prefer the VTID if present
  @VTID(117)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  com.exceljava.com4j.excel.Range getUsedRange();

  /**
   * <p>
   * Getter method for the COM property "UsedRange"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @DISPID(412) //= 0x19c. The runtime will prefer the VTID if present
  @VTID(117)
  com.exceljava.com4j.excel.Range getUsedRange(
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "HPageBreaks"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.HPageBreaks
   */

  @DISPID(1418) //= 0x58a. The runtime will prefer the VTID if present
  @VTID(118)
  com.exceljava.com4j.excel.HPageBreaks getHPageBreaks();


  /**
   * <p>
   * Getter method for the COM property "VPageBreaks"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.VPageBreaks
   */

  @DISPID(1419) //= 0x58b. The runtime will prefer the VTID if present
  @VTID(119)
  com.exceljava.com4j.excel.VPageBreaks getVPageBreaks();


  /**
   * <p>
   * Getter method for the COM property "QueryTables"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.QueryTables
   */

  @DISPID(1434) //= 0x59a. The runtime will prefer the VTID if present
  @VTID(120)
  com.exceljava.com4j.excel.QueryTables getQueryTables();


  /**
   * <p>
   * Getter method for the COM property "DisplayPageBreaks"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1435) //= 0x59b. The runtime will prefer the VTID if present
  @VTID(121)
  boolean getDisplayPageBreaks();


  /**
   * <p>
   * Setter method for the COM property "DisplayPageBreaks"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1435) //= 0x59b. The runtime will prefer the VTID if present
  @VTID(122)
  void setDisplayPageBreaks(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Comments"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Comments
   */

  @DISPID(575) //= 0x23f. The runtime will prefer the VTID if present
  @VTID(123)
  com.exceljava.com4j.excel.Comments getComments();


  /**
   * <p>
   * Getter method for the COM property "Hyperlinks"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Hyperlinks
   */

  @DISPID(1393) //= 0x571. The runtime will prefer the VTID if present
  @VTID(124)
  com.exceljava.com4j.excel.Hyperlinks getHyperlinks();


  /**
   */

  @DISPID(1436) //= 0x59c. The runtime will prefer the VTID if present
  @VTID(125)
  void clearCircles();


  /**
   */

  @DISPID(1437) //= 0x59d. The runtime will prefer the VTID if present
  @VTID(126)
  void circleInvalid();


  /**
   * <p>
   * Getter method for the COM property "_DisplayRightToLeft"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * get_DisplayRightToLeft(1033);
   * </code>
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(648) //= 0x288. The runtime will prefer the VTID if present
  @VTID(127)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  int get_DisplayRightToLeft();

  /**
   * <p>
   * Getter method for the COM property "_DisplayRightToLeft"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type int
   */

  @DISPID(648) //= 0x288. The runtime will prefer the VTID if present
  @VTID(127)
  int get_DisplayRightToLeft(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "_DisplayRightToLeft"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * set_DisplayRightToLeft(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(648) //= 0x288. The runtime will prefer the VTID if present
  @VTID(128)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void set_DisplayRightToLeft(
    int rhs);

  /**
   * <p>
   * Setter method for the COM property "_DisplayRightToLeft"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory int parameter.
   */

  @DISPID(648) //= 0x288. The runtime will prefer the VTID if present
  @VTID(128)
  void set_DisplayRightToLeft(
    @LCID int lcid,
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "AutoFilter"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.AutoFilter
   */

  @DISPID(793) //= 0x319. The runtime will prefer the VTID if present
  @VTID(129)
  com.exceljava.com4j.excel.AutoFilter getAutoFilter();


  /**
   * <p>
   * Getter method for the COM property "DisplayRightToLeft"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getDisplayRightToLeft(1033);
   * </code>
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1774) //= 0x6ee. The runtime will prefer the VTID if present
  @VTID(130)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  @ReturnValue(index=1)
  boolean getDisplayRightToLeft();

  /**
   * <p>
   * Getter method for the COM property "DisplayRightToLeft"
   * </p>
   * @param lcid Mandatory int parameter.
   * @return  Returns a value of type boolean
   */

  @DISPID(1774) //= 0x6ee. The runtime will prefer the VTID if present
  @VTID(130)
  boolean getDisplayRightToLeft(
    @LCID int lcid);


  /**
   * <p>
   * Setter method for the COM property "DisplayRightToLeft"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setDisplayRightToLeft(1033, rhs);
   * </code>
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1774) //= 0x6ee. The runtime will prefer the VTID if present
  @VTID(131)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void setDisplayRightToLeft(
    boolean rhs);

  /**
   * <p>
   * Setter method for the COM property "DisplayRightToLeft"
   * </p>
   * @param lcid Mandatory int parameter.
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1774) //= 0x6ee. The runtime will prefer the VTID if present
  @VTID(131)
  void setDisplayRightToLeft(
    @LCID int lcid,
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Scripts"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Scripts
   */

  @DISPID(1816) //= 0x718. The runtime will prefer the VTID if present
  @VTID(132)
  com.exceljava.com4j.office.Scripts getScripts();


  @VTID(132)
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
  @VTID(133)
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
  @VTID(133)
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
  @VTID(133)
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
  @VTID(133)
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
  @VTID(133)
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
  @VTID(133)
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
  @VTID(133)
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
  @VTID(133)
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
  @VTID(133)
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
  @VTID(133)
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
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter customDictionary is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreUppercase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter alwaysSuggest is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter spellLang is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreFinalYaa is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter spellScript is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _CheckSpelling(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(1817) //= 0x719. The runtime will prefer the VTID if present
  @VTID(134)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _CheckSpelling();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter ignoreUppercase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter alwaysSuggest is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter spellLang is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreFinalYaa is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter spellScript is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _CheckSpelling(customDictionary, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1817) //= 0x719. The runtime will prefer the VTID if present
  @VTID(134)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void _CheckSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter alwaysSuggest is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter spellLang is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreFinalYaa is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter spellScript is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _CheckSpelling(customDictionary, ignoreUppercase, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreUppercase Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1817) //= 0x719. The runtime will prefer the VTID if present
  @VTID(134)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void _CheckSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreUppercase);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter spellLang is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreFinalYaa is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter spellScript is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _CheckSpelling(customDictionary, ignoreUppercase, alwaysSuggest, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreUppercase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alwaysSuggest Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1817) //= 0x719. The runtime will prefer the VTID if present
  @VTID(134)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void _CheckSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreUppercase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object alwaysSuggest);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter ignoreFinalYaa is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter spellScript is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _CheckSpelling(customDictionary, ignoreUppercase, alwaysSuggest, spellLang, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreUppercase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alwaysSuggest Optional parameter. Default value is com4j.Variant.getMissing()
   * @param spellLang Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1817) //= 0x719. The runtime will prefer the VTID if present
  @VTID(134)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void _CheckSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreUppercase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object alwaysSuggest,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object spellLang);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter spellScript is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _CheckSpelling(customDictionary, ignoreUppercase, alwaysSuggest, spellLang, ignoreFinalYaa, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreUppercase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alwaysSuggest Optional parameter. Default value is com4j.Variant.getMissing()
   * @param spellLang Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreFinalYaa Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1817) //= 0x719. The runtime will prefer the VTID if present
  @VTID(134)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void _CheckSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreUppercase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object alwaysSuggest,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object spellLang,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreFinalYaa);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * _CheckSpelling(customDictionary, ignoreUppercase, alwaysSuggest, spellLang, ignoreFinalYaa, spellScript, 1033);
   * </code>
   * </p>
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreUppercase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alwaysSuggest Optional parameter. Default value is com4j.Variant.getMissing()
   * @param spellLang Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreFinalYaa Optional parameter. Default value is com4j.Variant.getMissing()
   * @param spellScript Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1817) //= 0x719. The runtime will prefer the VTID if present
  @VTID(134)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void _CheckSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreUppercase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object alwaysSuggest,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object spellLang,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreFinalYaa,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object spellScript);

  /**
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreUppercase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alwaysSuggest Optional parameter. Default value is com4j.Variant.getMissing()
   * @param spellLang Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreFinalYaa Optional parameter. Default value is com4j.Variant.getMissing()
   * @param spellScript Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(1817) //= 0x719. The runtime will prefer the VTID if present
  @VTID(134)
  void _CheckSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreUppercase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object alwaysSuggest,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object spellLang,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreFinalYaa,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object spellScript,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "Tab"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Tab
   */

  @DISPID(1041) //= 0x411. The runtime will prefer the VTID if present
  @VTID(135)
  com.exceljava.com4j.excel.Tab getTab();


  /**
   * <p>
   * Getter method for the COM property "MailEnvelope"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.IMsoEnvelopeVB
   */

  @DISPID(2021) //= 0x7e5. The runtime will prefer the VTID if present
  @VTID(136)
  com.exceljava.com4j.office.IMsoEnvelopeVB getMailEnvelope();


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
  @VTID(137)
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
  @VTID(137)
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
  @VTID(137)
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
  @VTID(137)
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
  @VTID(137)
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
  @VTID(137)
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
  @VTID(137)
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
  @VTID(137)
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
  @VTID(137)
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
  @VTID(137)
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
   * Getter method for the COM property "CustomProperties"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.CustomProperties
   */

  @DISPID(2030) //= 0x7ee. The runtime will prefer the VTID if present
  @VTID(138)
  com.exceljava.com4j.excel.CustomProperties getCustomProperties();


  /**
   * <p>
   * Getter method for the COM property "SmartTags"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SmartTags
   */

  @DISPID(2016) //= 0x7e0. The runtime will prefer the VTID if present
  @VTID(139)
  com.exceljava.com4j.excel.SmartTags getSmartTags();


  /**
   * <p>
   * Getter method for the COM property "Protection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Protection
   */

  @DISPID(176) //= 0xb0. The runtime will prefer the VTID if present
  @VTID(140)
  com.exceljava.com4j.excel.Protection getProtection();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter format is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter link is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter displayAsIcon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter noHTMLFormatting is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pasteSpecial(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(1928) //= 0x788. The runtime will prefer the VTID if present
  @VTID(141)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void pasteSpecial();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter link is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter displayAsIcon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter noHTMLFormatting is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pasteSpecial(format, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1928) //= 0x788. The runtime will prefer the VTID if present
  @VTID(141)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void pasteSpecial(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter displayAsIcon is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter noHTMLFormatting is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pasteSpecial(format, link, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1928) //= 0x788. The runtime will prefer the VTID if present
  @VTID(141)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void pasteSpecial(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter iconFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter noHTMLFormatting is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pasteSpecial(format, link, displayAsIcon, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1928) //= 0x788. The runtime will prefer the VTID if present
  @VTID(141)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
  void pasteSpecial(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayAsIcon);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter iconIndex is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter noHTMLFormatting is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pasteSpecial(format, link, displayAsIcon, iconFileName, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconFileName Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1928) //= 0x788. The runtime will prefer the VTID if present
  @VTID(141)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
  void pasteSpecial(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayAsIcon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconFileName);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter iconLabel is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter noHTMLFormatting is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pasteSpecial(format, link, displayAsIcon, iconFileName, iconIndex, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconIndex Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1928) //= 0x788. The runtime will prefer the VTID if present
  @VTID(141)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
  void pasteSpecial(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayAsIcon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconIndex);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter noHTMLFormatting is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pasteSpecial(format, link, displayAsIcon, iconFileName, iconIndex, iconLabel, com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconLabel Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1928) //= 0x788. The runtime will prefer the VTID if present
  @VTID(141)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
  void pasteSpecial(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
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
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pasteSpecial(format, link, displayAsIcon, iconFileName, iconIndex, iconLabel, noHTMLFormatting, 1033);
   * </code>
   * </p>
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconLabel Optional parameter. Default value is com4j.Variant.getMissing()
   * @param noHTMLFormatting Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1928) //= 0x788. The runtime will prefer the VTID if present
  @VTID(141)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void pasteSpecial(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayAsIcon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconLabel,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object noHTMLFormatting);

  /**
   * @param format Optional parameter. Default value is com4j.Variant.getMissing()
   * @param link Optional parameter. Default value is com4j.Variant.getMissing()
   * @param displayAsIcon Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconIndex Optional parameter. Default value is com4j.Variant.getMissing()
   * @param iconLabel Optional parameter. Default value is com4j.Variant.getMissing()
   * @param noHTMLFormatting Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(1928) //= 0x788. The runtime will prefer the VTID if present
  @VTID(141)
  void pasteSpecial(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object format,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object link,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object displayAsIcon,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconIndex,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object iconLabel,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object noHTMLFormatting,
    @LCID int lcid);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter password is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter drawingObjects is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter contents is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter scenarios is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter userInterfaceOnly is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFormattingCells is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFormattingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFormattingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingHyperlinks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowSorting is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFiltering is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowUsingPivotTables is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(142)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void protect();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter drawingObjects is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter contents is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter scenarios is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter userInterfaceOnly is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFormattingCells is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFormattingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFormattingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingHyperlinks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowSorting is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFiltering is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowUsingPivotTables is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(142)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter contents is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter scenarios is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter userInterfaceOnly is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFormattingCells is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFormattingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFormattingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingHyperlinks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowSorting is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFiltering is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowUsingPivotTables is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, drawingObjects, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(142)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter scenarios is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter userInterfaceOnly is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFormattingCells is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFormattingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFormattingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingHyperlinks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowSorting is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFiltering is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowUsingPivotTables is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, drawingObjects, contents, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(142)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter userInterfaceOnly is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFormattingCells is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFormattingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFormattingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingHyperlinks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowSorting is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFiltering is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowUsingPivotTables is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, drawingObjects, contents, scenarios, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scenarios Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(142)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scenarios);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter allowFormattingCells is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFormattingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFormattingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingHyperlinks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowSorting is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFiltering is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowUsingPivotTables is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, drawingObjects, contents, scenarios, userInterfaceOnly, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scenarios Optional parameter. Default value is com4j.Variant.getMissing()
   * @param userInterfaceOnly Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(142)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
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
     * <li>java.lang.Object parameter allowFormattingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFormattingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingHyperlinks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowSorting is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFiltering is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowUsingPivotTables is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scenarios Optional parameter. Default value is com4j.Variant.getMissing()
   * @param userInterfaceOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingCells Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(142)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scenarios,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object userInterfaceOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingCells);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter allowFormattingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingHyperlinks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowSorting is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFiltering is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowUsingPivotTables is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scenarios Optional parameter. Default value is com4j.Variant.getMissing()
   * @param userInterfaceOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingCells Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(142)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scenarios,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object userInterfaceOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingCells,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingColumns);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter allowInsertingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingHyperlinks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowSorting is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFiltering is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowUsingPivotTables is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scenarios Optional parameter. Default value is com4j.Variant.getMissing()
   * @param userInterfaceOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingCells Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingRows Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(142)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scenarios,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object userInterfaceOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingCells,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingRows);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter allowInsertingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowInsertingHyperlinks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowSorting is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFiltering is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowUsingPivotTables is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scenarios Optional parameter. Default value is com4j.Variant.getMissing()
   * @param userInterfaceOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingCells Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(142)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9, 10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scenarios,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object userInterfaceOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingCells,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingRows,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingColumns);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter allowInsertingHyperlinks is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowSorting is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFiltering is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowUsingPivotTables is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scenarios Optional parameter. Default value is com4j.Variant.getMissing()
   * @param userInterfaceOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingCells Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingRows Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(142)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, optParamIndex = {10, 11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scenarios,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object userInterfaceOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingCells,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingRows,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingRows);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter allowDeletingColumns is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowDeletingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowSorting is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFiltering is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowUsingPivotTables is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scenarios Optional parameter. Default value is com4j.Variant.getMissing()
   * @param userInterfaceOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingCells Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingHyperlinks Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(142)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10}, optParamIndex = {11, 12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scenarios,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object userInterfaceOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingCells,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingRows,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingRows,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingHyperlinks);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter allowDeletingRows is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowSorting is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFiltering is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowUsingPivotTables is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scenarios Optional parameter. Default value is com4j.Variant.getMissing()
   * @param userInterfaceOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingCells Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingHyperlinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowDeletingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(142)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}, optParamIndex = {12, 13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scenarios,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object userInterfaceOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingCells,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingRows,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingRows,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingHyperlinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowDeletingColumns);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter allowSorting is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowFiltering is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowUsingPivotTables is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scenarios Optional parameter. Default value is com4j.Variant.getMissing()
   * @param userInterfaceOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingCells Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingHyperlinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowDeletingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowDeletingRows Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(142)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}, optParamIndex = {13, 14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scenarios,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object userInterfaceOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingCells,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingRows,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingRows,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingHyperlinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowDeletingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowDeletingRows);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter allowFiltering is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter allowUsingPivotTables is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows, allowSorting, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scenarios Optional parameter. Default value is com4j.Variant.getMissing()
   * @param userInterfaceOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingCells Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingHyperlinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowDeletingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowDeletingRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowSorting Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(142)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13}, optParamIndex = {14, 15}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scenarios,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object userInterfaceOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingCells,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingRows,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingRows,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingHyperlinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowDeletingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowDeletingRows,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowSorting);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter allowUsingPivotTables is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * protect(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows, allowSorting, allowFiltering, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scenarios Optional parameter. Default value is com4j.Variant.getMissing()
   * @param userInterfaceOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingCells Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingHyperlinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowDeletingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowDeletingRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowSorting Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFiltering Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(142)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14}, optParamIndex = {15}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scenarios,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object userInterfaceOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingCells,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingRows,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingRows,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingHyperlinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowDeletingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowDeletingRows,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowSorting,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFiltering);

  /**
   * @param password Optional parameter. Default value is com4j.Variant.getMissing()
   * @param drawingObjects Optional parameter. Default value is com4j.Variant.getMissing()
   * @param contents Optional parameter. Default value is com4j.Variant.getMissing()
   * @param scenarios Optional parameter. Default value is com4j.Variant.getMissing()
   * @param userInterfaceOnly Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingCells Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFormattingRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowInsertingHyperlinks Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowDeletingColumns Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowDeletingRows Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowSorting Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowFiltering Optional parameter. Default value is com4j.Variant.getMissing()
   * @param allowUsingPivotTables Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2029) //= 0x7ed. The runtime will prefer the VTID if present
  @VTID(142)
  void protect(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object password,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object drawingObjects,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object contents,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object scenarios,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object userInterfaceOnly,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingCells,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFormattingRows,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingRows,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowInsertingHyperlinks,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowDeletingColumns,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowDeletingRows,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowSorting,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowFiltering,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object allowUsingPivotTables);


  /**
   * <p>
   * Getter method for the COM property "ListObjects"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObjects
   */

  @DISPID(2259) //= 0x8d3. The runtime will prefer the VTID if present
  @VTID(143)
  com.exceljava.com4j.excel.ListObjects getListObjects();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter selectionNamespaces is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter map is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * xmlDataQuery(xPath, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param xPath Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @DISPID(2260) //= 0x8d4. The runtime will prefer the VTID if present
  @VTID(144)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.excel.Range xmlDataQuery(
    java.lang.String xPath);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter map is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * xmlDataQuery(xPath, selectionNamespaces, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param xPath Mandatory java.lang.String parameter.
   * @param selectionNamespaces Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @DISPID(2260) //= 0x8d4. The runtime will prefer the VTID if present
  @VTID(144)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.excel.Range xmlDataQuery(
    java.lang.String xPath,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object selectionNamespaces);

  /**
   * @param xPath Mandatory java.lang.String parameter.
   * @param selectionNamespaces Optional parameter. Default value is com4j.Variant.getMissing()
   * @param map Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @DISPID(2260) //= 0x8d4. The runtime will prefer the VTID if present
  @VTID(144)
  com.exceljava.com4j.excel.Range xmlDataQuery(
    java.lang.String xPath,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object selectionNamespaces,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object map);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter selectionNamespaces is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter map is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * xmlMapQuery(xPath, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param xPath Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @DISPID(2263) //= 0x8d7. The runtime will prefer the VTID if present
  @VTID(145)
  @UseDefaultValues(paramIndexMapping = {0, 3}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.excel.Range xmlMapQuery(
    java.lang.String xPath);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter map is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * xmlMapQuery(xPath, selectionNamespaces, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param xPath Mandatory java.lang.String parameter.
   * @param selectionNamespaces Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @DISPID(2263) //= 0x8d7. The runtime will prefer the VTID if present
  @VTID(145)
  @UseDefaultValues(paramIndexMapping = {0, 1, 3}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=3)
  com.exceljava.com4j.excel.Range xmlMapQuery(
    java.lang.String xPath,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object selectionNamespaces);

  /**
   * @param xPath Mandatory java.lang.String parameter.
   * @param selectionNamespaces Optional parameter. Default value is com4j.Variant.getMissing()
   * @param map Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @DISPID(2263) //= 0x8d7. The runtime will prefer the VTID if present
  @VTID(145)
  com.exceljava.com4j.excel.Range xmlMapQuery(
    java.lang.String xPath,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object selectionNamespaces,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object map);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter from is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(146)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void printOut();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter to is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(146)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter copies is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(146)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter preview is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(146)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3, 4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter activePrinter is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(146)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004", "1033"})
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
     * <li>java.lang.Object parameter printToFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
   * </code>
   * </p>
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(146)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "80020004", "1033"})
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
     * <li>java.lang.Object parameter collate is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, printToFile, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
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
  @VTID(146)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "80020004", "1033"})
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
     * <li>java.lang.Object parameter prToFileName is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, printToFile, collate, com4j.Variant.getMissing(), com4j.Variant.getMissing(), 1033);
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
  @VTID(146)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8, 9}, javaType = {java.lang.Object.class, java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "80020004", "1033"})
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
     * <li>java.lang.Object parameter ignorePrintAreas is set to com4j.Variant.getMissing()</li><li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, printToFile, collate, prToFileName, com4j.Variant.getMissing(), 1033);
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
  @VTID(146)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8, 9}, javaType = {java.lang.Object.class, int.class}, nativeType = {NativeType.VARIANT, NativeType.Int32}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_I4}, literal = {"80020004", "1033"})
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
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter lcid is set to 1033</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * printOut(from, to, copies, preview, activePrinter, printToFile, collate, prToFileName, ignorePrintAreas, 1033);
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
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(146)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7, 8}, optParamIndex = {9}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"1033"})
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object prToFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas);

  /**
   * @param from Optional parameter. Default value is com4j.Variant.getMissing()
   * @param to Optional parameter. Default value is com4j.Variant.getMissing()
   * @param copies Optional parameter. Default value is com4j.Variant.getMissing()
   * @param preview Optional parameter. Default value is com4j.Variant.getMissing()
   * @param activePrinter Optional parameter. Default value is com4j.Variant.getMissing()
   * @param printToFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param collate Optional parameter. Default value is com4j.Variant.getMissing()
   * @param prToFileName Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignorePrintAreas Optional parameter. Default value is com4j.Variant.getMissing()
   * @param lcid Mandatory int parameter.
   */

  @DISPID(2361) //= 0x939. The runtime will prefer the VTID if present
  @VTID(146)
  void printOut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object from,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object to,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object copies,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object preview,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object activePrinter,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object printToFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object collate,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object prToFileName,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignorePrintAreas,
    @LCID int lcid);


  /**
   * <p>
   * Getter method for the COM property "EnableFormatConditionsCalculation"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(2511) //= 0x9cf. The runtime will prefer the VTID if present
  @VTID(147)
  boolean getEnableFormatConditionsCalculation();


  /**
   * <p>
   * Setter method for the COM property "EnableFormatConditionsCalculation"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2511) //= 0x9cf. The runtime will prefer the VTID if present
  @VTID(148)
  void setEnableFormatConditionsCalculation(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Sort"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Sort
   */

  @DISPID(880) //= 0x370. The runtime will prefer the VTID if present
  @VTID(149)
  com.exceljava.com4j.excel.Sort getSort();


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
  @VTID(150)
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
  @VTID(150)
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
  @VTID(150)
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
  @VTID(150)
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
  @VTID(150)
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
  @VTID(150)
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
  @VTID(150)
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
  @VTID(150)
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
  @VTID(150)
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
   * Getter method for the COM property "PrintedCommentPages"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(2857) //= 0xb29. The runtime will prefer the VTID if present
  @VTID(151)
  int getPrintedCommentPages();


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
  @VTID(152)
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
  @VTID(152)
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
  @VTID(152)
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
  @VTID(152)
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
  @VTID(152)
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
  @VTID(152)
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
  @VTID(152)
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
  @VTID(152)
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
  @VTID(152)
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
  @VTID(152)
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
  @VTID(153)
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
  @VTID(153)
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
  @VTID(153)
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
  @VTID(153)
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
  @VTID(153)
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
  @VTID(153)
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
  @VTID(153)
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
  @VTID(153)
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
  @VTID(153)
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
  @VTID(153)
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
